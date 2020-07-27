VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBalAdjEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Adjustment Entry"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   2175
   ClientWidth     =   12210
   Icon            =   "frmBalAdjEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpLongInteger fpAcct 
      Height          =   324
      Left            =   4176
      TabIndex        =   0
      Top             =   1992
      Width           =   1872
      _Version        =   196608
      _ExtentX        =   3302
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
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   0
      Text            =   ""
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5:03 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "1/2/2008"
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   9384
      TabIndex        =   1
      Top             =   1800
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   1
      Left            =   9384
      TabIndex        =   2
      Top             =   2112
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   2
      Left            =   9384
      TabIndex        =   3
      Top             =   2424
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   3
      Left            =   9384
      TabIndex        =   4
      Top             =   2736
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   4
      Left            =   9384
      TabIndex        =   5
      Top             =   3048
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   5
      Left            =   9384
      TabIndex        =   6
      Top             =   3360
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   6
      Left            =   9384
      TabIndex        =   7
      Top             =   3672
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   7
      Left            =   9384
      TabIndex        =   8
      Top             =   3984
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   8
      Left            =   9384
      TabIndex        =   9
      Top             =   4296
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   9
      Left            =   9384
      TabIndex        =   10
      Top             =   4608
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   10
      Left            =   9384
      TabIndex        =   11
      Top             =   4920
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   11
      Left            =   9384
      TabIndex        =   12
      Top             =   5232
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   12
      Left            =   9384
      TabIndex        =   13
      Top             =   5544
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   13
      Left            =   9384
      TabIndex        =   14
      Top             =   5856
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   14
      Left            =   9384
      TabIndex        =   15
      Top             =   6168
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8772
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBalAdjEntry.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   384
      Left            =   10080
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBalAdjEntry.frx":0AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdHist 
      Height          =   384
      Left            =   1080
      TabIndex        =   20
      Top             =   7440
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
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
      ButtonDesigner  =   "frmBalAdjEntry.frx":0C82
   End
   Begin EditLib.fpText fpCustRecNo 
      Height          =   324
      Left            =   504
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   648
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      NoSpecialKeys   =   3
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
      Text            =   "fpText1"
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdFind 
      Height          =   390
      Left            =   2385
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmBalAdjEntry.frx":0E5D
   End
   Begin EditLib.fpText fpstatus 
      Height          =   300
      Left            =   432
      TabIndex        =   31
      Top             =   48
      Visible         =   0   'False
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
      _ExtentY        =   529
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
      ControlType     =   1
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
   Begin EditLib.fpText fptaxexmpt 
      Height          =   300
      Left            =   432
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   1884
      _Version        =   196608
      _ExtentX        =   3323
      _ExtentY        =   529
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
      ControlType     =   1
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
   Begin EditLib.fpDoubleSingle fpPrevAmt 
      Height          =   324
      Left            =   9408
      TabIndex        =   17
      Top             =   6912
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpCurrAmt 
      Height          =   324
      Left            =   9408
      TabIndex        =   16
      Top             =   6576
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
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
   Begin EditLib.fpDoubleSingle fpDepAmt 
      Height          =   324
      Left            =   2136
      TabIndex        =   53
      Top             =   4944
      Width           =   1596
      _Version        =   196608
      _ExtentX        =   2815
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
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
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
      DecimalPlaces   =   2
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
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Balance:"
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
      Left            =   6840
      TabIndex        =   52
      Top             =   6936
      Width           =   2472
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance:"
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
      Left            =   6960
      TabIndex        =   51
      Top             =   6600
      Width           =   2268
   End
   Begin VB.Label fptxtCity 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2160
      TabIndex        =   50
      Top             =   3216
      Width           =   3924
   End
   Begin VB.Label fptxtAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2160
      TabIndex        =   49
      Top             =   2904
      Width           =   3924
   End
   Begin VB.Label fptxtName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2160
      TabIndex        =   48
      Top             =   2592
      Width           =   3924
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   14
      Left            =   6408
      TabIndex        =   33
      Top             =   6168
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   13
      Left            =   6408
      TabIndex        =   34
      Top             =   5868
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   10
      Left            =   6408
      TabIndex        =   37
      Top             =   4920
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   9
      Left            =   6408
      TabIndex        =   38
      Top             =   4620
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   7
      Left            =   6408
      TabIndex        =   40
      Top             =   3996
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   5
      Left            =   6408
      TabIndex        =   42
      Top             =   3372
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   4
      Left            =   6408
      TabIndex        =   43
      Top             =   3072
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   1
      Left            =   6408
      TabIndex        =   46
      Top             =   2100
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   0
      Left            =   6408
      TabIndex        =   47
      Top             =   1800
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   2
      Left            =   6408
      TabIndex        =   45
      Top             =   2424
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   3
      Left            =   6408
      TabIndex        =   44
      Top             =   2748
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   6
      Left            =   6408
      TabIndex        =   41
      Top             =   3696
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   8
      Left            =   6408
      TabIndex        =   39
      Top             =   4320
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   11
      Left            =   6408
      TabIndex        =   36
      Top             =   5244
      Width           =   2940
   End
   Begin VB.Label fpRevSource 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   12
      Left            =   6408
      TabIndex        =   35
      Top             =   5568
      Width           =   2940
   End
   Begin VB.Shape Shape3 
      Height          =   612
      Left            =   768
      Top             =   7320
      Width           =   10716
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   5964
      Left            =   774
      Top             =   1344
      Width           =   10668
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   6384
      X2              =   11304
      Y1              =   6528
      Y2              =   6528
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Revenue Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6384
      TabIndex        =   29
      Top             =   1488
      Width           =   2964
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Adjustment Entry"
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
      TabIndex        =   28
      Top             =   516
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2580
      Top             =   432
      Width           =   7020
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account Number:"
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
      Left            =   1200
      TabIndex        =   27
      Top             =   2040
      Width           =   2856
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   1044
      TabIndex        =   26
      Top             =   2640
      Width           =   972
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep Amt:"
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
      Left            =   840
      TabIndex        =   25
      Top             =   4992
      Width           =   1188
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   792
      TabIndex        =   24
      Top             =   3060
      Width           =   1248
   End
   Begin VB.Line Line3 
      X1              =   6372
      X2              =   6372
      Y1              =   1368
      Y2              =   7272
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Balance Amts"
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
      Left            =   9384
      TabIndex        =   23
      Top             =   1488
      Width           =   1788
   End
   Begin VB.Line Line4 
      X1              =   9360
      X2              =   9360
      Y1              =   1488
      Y2              =   7296
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2592
      Top             =   312
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBalAdjEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CustAcct As Long, TransRec As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean, PayListCnt As Long
Dim DistArray() As DistArrayType
Dim fromform As Form, toform As Form, codeopt As Integer, noreset As Boolean
Dim RevText$(1 To MaxRevsCnt)
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub
Private Sub Form_Activate()
  If Val(fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    fpAcct = fpCustRecNo
    GetCustinfo
    DoEvents
  End If
End Sub

Private Sub cmdExit_Click()
'  ChkEmptyAcct
'  noreset = True
'  Chk4Change
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    CheckInfo
'    If ChkOKFlag Then
'      fpCmdSave_Click
'    Else
'      Exit Sub
'    End If
'  End If
  CustAcct = 0
  fpCustRecNo = 0
  BeenDone = False
'  If codeopt = 1 Then
'    ActivateControls frmCustEditLookUP
'  ElseIf codeopt = 2 Then
'    ActivateControls frmDisplayList
'  End If
'  If codeopt = 0 Then
    Load frmUBEditMenu
    DoEvents
    frmUBEditMenu.Show
'  End If
  Erase RevText$
  'UBLog "OUT: UBBalAdj"
  Unload Me
  DoEvents
End Sub

Private Sub fpAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn, vbKeyDown, vbKeyUp, vbKeyTab
      KeyCode = 0
    If Len(fpAcct) > 0 Then
        fpDepAmt.SetFocus
    End If
  End Select
End Sub
Private Sub ChkEmptyAcct()
  If Len(fpAcct) <= 0 Then
    ClearScn
  End If
End Sub

'Private Sub fpAmtPaid_LostFocus(Index As Integer)
'  CalcBALFlds
'End Sub
'
'Private Sub fpAmtPaid_ChangeMode(Index As Integer, EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub fpAmtPaid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'  Dim x As Integer
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
'    If Index < MaxRevsCnt Then
'     For x = Index To (MaxRevsCnt - 1)
'      If fpAmtPaid(x + 1).Enabled Then
'        fpAmtPaid(x + 1).SetFocus
'        Exit For
'      Else
'        fpCmdSave.SetFocus
'        Exit For
'      End If
'     Next
'    End If
'  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
'    If Index > 0 Then
'     For x = Index To (MaxRevsCnt - 1)
'      If fpAmtPaid(x - 1).Enabled Then
'        fpAmtPaid(x - 1).SetFocus
'        Exit For
'      Else
'        fptxtDesc.SetFocus
'      End If
'     Next
'    End If
'  End If
'
'End Sub

Private Sub fpCurrAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

'Private Sub fpCashAmt_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    If fpChkAmt.Enabled Then
'      fpChkAmt.SetFocus
'    Else
'      fptxtDesc.SetFocus
'    End If
'  End If
'End Sub


Private Sub fpPrevAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

'Private Sub fpChkAmt_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'    fptxtDesc.SetFocus
'  End If
'End Sub
'
'Private Sub fpCmdDist_Click()
'If Len(fpAcct) > 0 Then
'  If fpTotReceived > 0 Then
'    TempAmtRecv = fpTotReceived
'    Autodist
'  End If
'End If
'End Sub
'Private Sub Chk4Change()
'  Dim cntout As Integer, cnt As Integer
'  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
'
'  Dim PayFileName As String, UBPayRecLen As Integer
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'
'  UBPayRecLen = Len(UBPaymentRec(1))
'  If Len(fpAcct) > 0 Then
'  NumofRevs = MaxRevsCnt
'  cntout = 0
'  Answer = 0
'  If EditFlag = True Then
'    PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
'    ListFile = FreeFile
'    Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
'    Get ListFile, PayListRec&, UBPaymentRec(1)
'    Close ListFile
'    If fpTAmtOwed <> UBPaymentRec(1).AMTOWED Then cntout = cntout + 1
'    If txtPaymentDate <> Num2Date(UBPaymentRec(1).PAYDATE) Then cntout = cntout + 1
'    If QPTrim$(fptxtDesc) <> QPTrim$(UBPaymentRec(1).Desc) Then cntout = cntout + 1
'    For cnt = 1 To NumofRevs
'      If fpAmtOwed(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 Then cntout = cntout + 1
'      If fpAmtPaid(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTPD1 Then cntout = cntout + 1
'    Next
'    Select Case QPTrim(UBPaymentRec(1).TENDERTY)
'      Case "Cash":
'        If fpcboTenderType.ListIndex <> 0 Then cntout = cntout + 1
'      Case "Check":
'        If fpcboTenderType.ListIndex <> 1 Then cntout = cntout + 1
'      Case "Cash & Check":
'        If fpcboTenderType.ListIndex <> 2 Then cntout = cntout + 1
'      Case "Charge":
'        If fpcboTenderType.ListIndex <> 3 Then cntout = cntout + 1
'    End Select
'    If fpCashAmt <> UBPaymentRec(1).CASHAMT Then cntout = cntout + 1
'    If fpChkAmt <> UBPaymentRec(1).CHKAMT Then cntout = cntout + 1
'    If fpTotReceived <> UBPaymentRec(1).AMTRECD Then cntout = cntout + 1
'    If fpChangeDue <> UBPaymentRec(1).Change Then cntout = cntout + 1
'  Else
'
'    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
'  End If
'  If cntout > 0 Then
'    frmChangedWarning.Show vbModal, Me
'    Select Case SaveFlag
'    Case False
'      Answer = 3
'    Case True
'      Answer = 2
'    Case 1
'      Answer = 1
'    End Select
'  Else
'    Answer = 0
'  End If
'  End If
'End Sub
'Private Sub Chk4OKforNew()
'  Dim FntSize As Integer
'  Dim cntout As Integer, cnt As Integer
'  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
'
'  Dim PayFileName As String, UBPayRecLen As Integer
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'
'  UBPayRecLen = Len(UBPaymentRec(1))
'  If Len(fpAcct) > 0 Then
'  NumofRevs = MaxRevsCnt
'  cntout = 0
'  Answer = 0
'  If EditFlag = True Then
'    PayFileName$ = UBPath$ + "UBPAY" + Oper$ + ".DAT"
'    ListFile = FreeFile
'    Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
'    Get ListFile, PayListRec&, UBPaymentRec(1)
'    Close ListFile
'    If fpTAmtOwed <> UBPaymentRec(1).AMTOWED Then cntout = cntout + 1
'    If txtPaymentDate <> Num2Date(UBPaymentRec(1).PAYDATE) Then cntout = cntout + 1
'    If QPTrim$(fptxtDesc) <> QPTrim$(UBPaymentRec(1).Desc) Then cntout = cntout + 1
'    For cnt = 1 To NumofRevs
'      If fpAmtOwed(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 Then cntout = cntout + 1
'      If fpAmtPaid(cnt - 1) <> UBPaymentRec(1).PaidOwed(cnt).AMTPD1 Then cntout = cntout + 1
'    Next
'    Select Case QPTrim(UBPaymentRec(1).TENDERTY)
'      Case "Cash":
'        If fpcboTenderType.ListIndex <> 0 Then cntout = cntout + 1
'      Case "Check":
'        If fpcboTenderType.ListIndex <> 1 Then cntout = cntout + 1
'      Case "Cash & Check":
'        If fpcboTenderType.ListIndex <> 2 Then cntout = cntout + 1
'      Case "Charge":
'        If fpcboTenderType.ListIndex <> 3 Then cntout = cntout + 1
'    End Select
'    If fpCashAmt <> UBPaymentRec(1).CASHAMT Then cntout = cntout + 1
'    If fpChkAmt <> UBPaymentRec(1).CHKAMT Then cntout = cntout + 1
'    If fpTotReceived <> UBPaymentRec(1).AMTRECD Then cntout = cntout + 1
'    If fpChangeDue <> UBPaymentRec(1).Change Then cntout = cntout + 1
'  Else
'
'    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
'  End If
'  If cntout > 0 Then
'    ReDim MsgText(0 To 5) As String
'    FntSize = frmMsgDialog.Label(1).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "WARNING:Payment In Progress"
'    MsgText(1) = ""
'    MsgText(2) = "Do You Want to Abandon this Payment?"
'    MsgText(3) = "Ok to Abandon,"
'    MsgText(4) = "Cancel to Remain on Current Payment."
'    MsgText(5) = ""
'    If GetOKorNot(MsgText()) Then
'     UBLog "USER WANTS TO Abandon"
'     Answer = 2
'    Else
'     UBLog "USER Canceled"
'     Answer = 1
'    End If
'  Else
'    Answer = 0
'  End If
'  End If
'End Sub
'
'Private Sub fpcmdDrawer_Click()
'  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
'  On Local Error Resume Next
'  If RecpDef = 99 Then Exit Sub
'  'RecPort = GetDEFPort%
'  Port$ = QPTrim$(RecpPort)
'
'   'DPName = QPTrim(Printers(2).DeviceName) 'RecpPort).DeviceName)
'   ' DefPrinter = QPTrim(Printers(2).Port) 'RecpPort).Port)
'  'Added this to allow for winxp network printer port names of ne00:, etc.
'  'the device name worked so use that instead, but only for network printers.
''    If InStr(1, DPName, "\\", vbTextCompare) Then
''      Port$ = DPName
''      'frmViewPrint.PrintWSet DPName, Copies
''    Else
''      Port$ = DefPrinter
''      'frmViewPrint.PrintWSet DefPrinter, Copies
''    End If
''    If vbKeyDown = vbKeyEscape Then
''      Printer.KillDoc
''    End If
'
'
'  UBLog "Oper: " + Oper$ + "Util Pay-Open Drawer"
'  PortFile = FreeFile
'  Open Port$ For Output As #PortFile
'  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
'  Print #PortFile, Chr$(7)
'  Close PortFile
'End Sub

Private Sub fpcmdFind_Click()
  'Chk4OKforNew
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    'continue on
'  End If
  ClearScn
  frmCustEditLookUP.Caption = "Payment Customer Find"
  frmCustEditLookUP.Label1.Caption = "Payment Customer Find"
  frmCustEditLookUP.Wheretogo frmBalAdjEntry, frmBalAdjEntry
  Unload Me
  DoEvents
  frmCustEditLookUP.Show
  DoEvents
End Sub
Private Sub fpCmdHist_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If TransRec > 0 Then
   ' DeActivateControls Me
    DisplayCustTransList CustAcct&
   ' ActivateControls Me
  Else
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "There are NO transactions to display."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If
End Sub

'Private Sub fpCmdInfo_Click()
'If Len(fpAcct) > 0 Then
'  If Len(fpCustRecNo) > 0 Then
'    'DeActivateControls Me
'    frmInfo.Label1 = "Loading. . ."
'    frmInfo.Show
'    DoEvents
'    'here
'    frmRptCustInq.fpCustRecNo = Me.fpCustRecNo
'    'frmRptCustInq.Wheretogo frmPaymentEntry, frmRptCustInq, 0
'    'Load frmRptCustInq
'    frmRptCustInq.Show
'    DoEvents
'    Unload frmInfo
'  End If
'End If
'End Sub
Private Sub fpCmdSave_Click()
 ' ChkEmptyAcct
  DoEvents
  If Len(fpAcct) <= 0 Then
    MsgBox "Invalid Account Information.", vbOKOnly, "Invalid Entry"
    Exit Sub
  End If
 ' CalcBALFlds
 ' CheckInfo
 ' If ChkOKFlag Then
    'DeActivateControls Me
'    frmPrintReceipt.Show 1
'    If SavePay = True Then
      SaveTransaction
    
'      If PrnRecp = True Then
'        PrintReceipt
'      End If
    
      MsgBox "Transaction Complete.", vbOKOnly, "Complete"
      ClearScn
      
'    End If
    
'      CustAcct = 0
'      fpCustRecNo = 0
'      BeenDone = False
'      If codeopt = 1 Then
'        ActivateControls frmCustEditLookUP
'      ElseIf codeopt = 2 Then
'        ActivateControls frmDisplayList
'      End If
'      If codeopt = 0 Then
'        Load frmUBPaymentMenu
'        DoEvents
'        frmUBPaymentMenu.Show
'      End If
'
'      UBLog "OUT: UTIL Payment" + " Oper:" + Oper$
'      Unload Me
'      DoEvents
'    ActivateControls Me
'  End If

End Sub
'Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
'    KeyCode = 0
'    fpAmount(0).SetFocus
'  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
'    fpCmdSave.SetFocus
'  End If
'End Sub

'Private Sub fpcmdCash_Click()
'If Len(fpAcct) > 0 Then
'  fpcboTenderType.ListIndex = 0
'  fpChkAmt.Enabled = False
'  fpCashAmt.Enabled = True
'  fpChkAmt = 0
'  fpCashAmt = fpTAmtOwed.DoubleValue
'  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
'  If fpTotReceived > 0 Then
'    TempAmtRecv = fpTotReceived
'    Autodist
'  End If
'  fptxtDesc.SetFocus
'End If
'End Sub

'Private Sub fpcmdCheck_Click()
'If Len(fpAcct) > 0 Then
'  fpcboTenderType.ListIndex = 1
'  fpCashAmt.Enabled = False
'  fpChkAmt.Enabled = True
'  fpCashAmt = 0
'  fpChkAmt = fpTAmtOwed.DoubleValue
'  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
'  If fpTotReceived > 0 Then
'    TempAmtRecv = fpTotReceived
'    Autodist
'  End If
'  fptxtDesc.SetFocus
'End If
'End Sub
'Private Sub fpCmdCharge_Click()
'If Len(fpAcct) > 0 Then
'  fpcboTenderType.ListIndex = 3
'  fpCashAmt.Enabled = False
'  fpChkAmt.Enabled = True
'  fpCashAmt = 0
'  fpChkAmt = fpTAmtOwed.DoubleValue
'  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
'  If fpTotReceived > 0 Then
'    TempAmtRecv = fpTotReceived
'    Autodist
'  End If
'  fpChangeDue.Enabled = False
'  fptxtDesc.SetFocus
'End If
'End Sub

'Private Sub fpCashAmt_LostFocus()
'fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
'If fpTotReceived > 0 Then
'  If fpcboTenderType.ListIndex <> 3 Then
'    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
'  End If
'End If
'End Sub

'Private Sub fpChkAmt_LostFocus()
'fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
'If fpTotReceived.DoubleValue > 0 Then
'  If fpcboTenderType.ListIndex <> 3 Then
'    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
'  End If
'End If
'End Sub
'Private Sub fpcboTenderType_DropDown()
'  ClrAmts
'End Sub
'
'Private Sub fpcboTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboTenderType.ListDown = True
'    'ClrAmts
'    KeyCode = 0
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcboTenderType.ListIndex = -1
'    fpcboTenderType.Action = ActionClearSearchBuffer
'    'ClrAmts
'  End If
'  If fpcboTenderType.ListDown <> True Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'      If fpCashAmt.Enabled = True Then
'        fpCashAmt.SetFocus
'      Else
'        fpChkAmt.SetFocus
'      End If
'        KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        fpAcct.SetFocus
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'  DoEvents
'End Sub
'Private Sub ClrAmts()
'  Dim cnt As Integer
'
'  fpCurrAmt = 0
'  fpPrevAmt = 0
'  For cnt = 1 To 15
'    fpAmtPaid(cnt - 1) = 0
'  Next
'  fpTotPaid = 0
'  fpTotReceived = 0
'End Sub
'Private Sub fpcboTenderType_SelChange(ItemIndex As Long)
'  If BeenDone Then
'    fixamts
' End If
'End Sub

'Private Sub fixamts()
'
'  fpcboTenderType.Action = ActionClearSearchBuffer
'  If noreset = False Then
'
'    If fpcboTenderType.ListIndex = 0 Then
'      fpCashAmt.Enabled = True
'      fpChkAmt.Enabled = False
'      'ClrAmts
'     ' fpCashAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 1 Then
'      fpCashAmt.Enabled = False
'      fpChkAmt.Enabled = True
'      'ClrAmts
'     ' fpChkAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 2 Then
'      fpCashAmt.Enabled = True
'      fpChkAmt.Enabled = True
'     ' ClrAmts
'     'fpCashAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 3 Then
'      fpCashAmt.Enabled = False
'      fpChkAmt.Enabled = True
'      fpChangeDue.Enabled = False
'     ' ClrAmts
'      'fpChkAmt.SetFocus
''    ElseIf fpcboTenderType.ListIndex = -1 Then
''      MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
''      fpcboTenderType.SetFocus
'    End If
'  End If
'  DoEvents
'  noreset = False
'End Sub
Private Sub fpAcct_LostFocus()
'Dim Acct As Long
'    Acct = fpAcct
'    If Acct > 0 Then
'      If Acct > GetTaxCustCnt Then
'        MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      ElseIf IsCustDeleted(Acct) Then
'        MsgBox "Deleted Account.", vbOKOnly, "Deleted Account"
'        fplngAcct.SetFocus
'        Exit Sub
'      Else
'       'If DoesCustOwe(Acct) Then
'          Cust2Screen (Acct)
'       ' Else
'       '   MsgBox "This Customer Does Not Owe A Balance.", vbOKOnly, "No Balance"
'      End If
'    Else
'      MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
'      fplngAcct.SetFocus
'      Exit Sub
'    End If
  fpCustRecNo = fpAcct
  'If Val(fpCustRecNo) > 0 Then
    GetCustinfo
  'Else
    'MsgBox "NO", vbOKOnly
  '  fpAcct.SetFocus
  'End If
End Sub

Private Sub fpDepAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpDepAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpAmtPaid(0).SetFocus
  End If
End Sub



Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PaymentEntry by " '+ PWUser$ + " operator-" + Oper$
       ' CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF4:
      KeyCode = 0
      DoEvents
      If fpCmdHist.Enabled Then
        Call fpCmdHist_Click
      End If
    Case vbKeyF7:
      KeyCode = 0
      DoEvents
      If fpcmdFind.Enabled Then
        Call fpcmdFind_Click
      End If
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdSave.Enabled Then
        Call fpCmdSave_Click
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  noreset = False
  TransRec = 0
  LoadRevs
  UBLog " IN UBBalADJ"
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
  ' Me.Visible = False
    Temp_Class.ResizeControls Me
  '  Me.Visible = True
  '  Me.SetFocus
  End If
  DoEvents
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub
Private Sub LoadRevs()
  Dim NumofRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, OutOfOrder As Boolean, x As Integer
  Dim tmp As DistArrayType
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  'ReDim RevText$(1 To MaxRevsCnt)
  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType
 ' RecpPort = Val(UBSetUpRec(1).RecpPort)
'  If RecpPort < 1 Or RecpPort > 2 Then
'    RecpPort = 1
'  End If

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        tmp = DistArray(x)
        DistArray(x) = DistArray(x + 1)
        DistArray(x + 1) = tmp
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

'  If NumOfRevs < MaxRevsCnt Then
'    ReDim Preserve RevText$(1 To NumOfRevs)
'  End If

  For RevCnt = 1 To NumofRevs
    fpRevSource(RevCnt - 1).Caption = RevText$(RevCnt)
  Next
'  For InvRev = NumofRevs To 14
'    fpRevSource(InvRev).Enabled = False
'    fpRevSource(InvRev).Visible = False
'    fpAmtPaid(InvRev).Enabled = False
'    fpAmtPaid(InvRev).Visible = False
'    fpAmtOwed(InvRev).Enabled = False
'    fpAmtOwed(InvRev).Visible = False
'  Next

End Sub
Private Sub GetCustinfo()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, cnt As Integer, TotalBalance As Double
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  NumofRevs = MaxRevsCnt
'  If uselook = True Then
'    Unload frmCustEditLookUP
'    Unload frmDisplayList
'    uselook = False
'  End If
  If fpCustRecNo <> "" Then
    CustAcct = fpCustRecNo
  Else
    'MsgBox "You Must Enter An Account Number.", vbOKOnly, "Invalid Account"
    fpAcct.SetFocus
    Exit Sub
  End If

  NumOfCustRecs& = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
    UBLog "ERROR: Invalid Account:" + Str$(CustAcct&) '+ " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
  
  If IsDeleted(CustAcct&) Then
    UBLog "ERROR: Deleted Account:" + Str$(CustAcct&) ' + " Oper:" + Oper$
    CustAcct& = 0
    'LabelDel.Caption = "Deleted Account!"
    'LabelDel.Visible = True
    GoTo SkiptoHere
  End If
 ' GoSub ClearForm
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  'FOR Cnt = 1 TO 15
  '  UBCustRec(1).CurrRevAmts(Cnt) = 0
  'NEXT
  'PUT CustFile, CUSTACCT&, UBCustRec(1)
  Close CustFile
  If UBCustRec(1).LastTrans > 0 Then
    TransRec = UBCustRec(1).LastTrans
  End If

  UBLog "Entering BAlADJ for Account:" + Str$(CustAcct&)
    For cnt = 1 To NumofRevs
      fpAmtPaid(cnt - 1) = UBCustRec(1).CurrRevAmts(cnt)
    Next
    fpCurrAmt = UBCustRec(1).CurrBalance
    fpPrevAmt = UBCustRec(1).PrevBalance
    fptxtName.Caption = UBCustRec(1).CustName
    fptxtAddress.Caption = UBCustRec(1).ADDR1
    fpstatus = UBCustRec(1).Status
    fpDepAmt = UBCustRec(1).DepositAmt
  
  CustAcct& = Val(fpCustRecNo)
  fptxtCity.Caption = UBCustRec(1).CITY
  fpDepAmt.SetFocus
  Exit Sub
SkiptoHere:
  BeenDone = True
  frmLookupError.Label.Caption = "Invalid Account Number"
  frmLookupError.Label1.Caption = "Please Enter A Valid Account Number."
  frmLookupError.Show 1
  ClearScn
  
End Sub
Private Sub ClearScn()
  Dim cnt As Integer
  BeenDone = False
  fpCustRecNo = 0
  fpAcct.Enabled = True
  fpAcct = ""
  'LabelDel.Visible = False
  'fpCmdTranHist.Enabled = False
  fpstatus = ""
  fptaxexmpt = ""
  TransRec = 0
  fptxtName.Caption = ""
  fptxtAddress.Caption = ""
  fptxtCity.Caption = ""
  fpDepAmt = 0
  fpCustRecNo = 0
  fpCurrAmt = 0
  fpPrevAmt = 0
  For cnt = 1 To 15
    fpAmtPaid(cnt - 1) = 0
  Next
  fpAcct.SetFocus
End Sub

'Private Sub CalcBALFlds()
'  Dim TOwd As Double, cnt As Integer, TPay As Double
'  TOwd# = 0
'  TPay# = 0
'  For cnt = 1 To MaxRevsCnt
'    TOwd# = Round#(TOwd# + fpAmtOwed(cnt - 1).DoubleValue)
'    'TCur# = Round#(TCur# + fpCurrent(cnt - 1).DoubleValue)
'    'fpActual(cnt - 1) = Round#(fpCurrent(cnt - 1).DoubleValue - fpAmount(cnt - 1).DoubleValue)
'    'fpActual(cnt - 1) = 0
'    TPay# = Round#(TPay# + fpAmtPaid(cnt - 1).DoubleValue)
'  Next
'  fpTotOwed = TOwd#
'  fpTotPaid = TPay#
'  If fpTotReceived > 0 Then
'    If fpcboTenderType.ListIndex <> 3 Then
'      fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
'    End If
'  End If
'End Sub
    
Private Sub SaveTransaction()
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim CustFile As Integer, cnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  
  For cnt = 1 To 15
    UBCustRec(1).CurrRevAmts(cnt) = fpAmtPaid(cnt - 1)
  Next
  UBCustRec(1).CurrBalance = fpCurrAmt
  UBCustRec(1).PrevBalance = fpPrevAmt
  UBCustRec(1).DepositAmt = fpDepAmt
  Put CustFile, CustAcct&, UBCustRec(1)
  Close
  UBLog "Updated Account:" + Str$(CustAcct&)
  

  'ClearScn
End Sub

'Private Sub CheckInfo()
'  Dim TestDate As Integer, TestAmt As Double
'  TestAmt = 0
'  ChkOKFlag = True
'  If fpTotReceived.DoubleValue <= 0 Or fpTotPaid.DoubleValue <= 0 Then
'    ChkOKFlag = False
'    MsgBox "Invalid Amount. The Total Received Should NOT Be ZERO.", vbOKOnly, "Request Canceled."
'    GoTo BadDate
'  End If
'  If fpChangeDue.DoubleValue >= 0 Then
'    TestAmt = Round#(fpTotReceived.DoubleValue - fpChangeDue.DoubleValue)
'    If TestAmt <> fpTotPaid Then '.DoubleValue Then
'      ChkOKFlag = False
'      MsgBox "The Total Received does NOT equal the Total Paid.", vbOKOnly, "Request Canceled."
'      GoTo BadDate
'    End If
'  Else
'    ChkOKFlag = False
'    MsgBox "The Amount Distributed May Not Be More Than Amount Received.", vbOKOnly, "Request Canceled."
'    GoTo BadDate
'  End If
'  Exit Sub
'BadDate:
'  Exit Sub
'End Sub
