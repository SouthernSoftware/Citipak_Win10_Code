VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Begin VB.Form frmEmpHistRptSplash 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9996
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   13728
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9996
   ScaleWidth      =   13728
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtSummary 
      Height          =   372
      Left            =   6384
      TabIndex        =   13
      Top             =   5760
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
   Begin EditLib.fpText fptxtEndDate 
      Height          =   372
      Left            =   6384
      TabIndex        =   12
      Top             =   5160
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
   Begin EditLib.fpText fptxtStartDate 
      Height          =   372
      Left            =   6384
      TabIndex        =   11
      Top             =   4560
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - "
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
   Begin EditLib.fpText fptxtLastEmpNo 
      Height          =   372
      Left            =   6384
      TabIndex        =   10
      Top             =   3960
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
   Begin EditLib.fpText fptxtFirstEmpNo 
      Height          =   372
      Left            =   6384
      TabIndex        =   9
      Top             =   3360
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
   Begin VB.CommandButton cmdProcess 
      Caption         =   "F10  &Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   7218
      TabIndex        =   8
      Top             =   6480
      Width           =   2292
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "ESC  &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4098
      TabIndex        =   7
      Top             =   6480
      Width           =   2412
   End
   Begin EditLib.fpText fpText7 
      Height          =   372
      Left            =   3984
      TabIndex        =   6
      Top             =   5760
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
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
      BackColor       =   -2147483638
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
      BorderStyle     =   1
      BorderColor     =   -2147483638
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
      AlignTextV      =   1
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
      Text            =   "Summaries Only:"
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
   Begin EditLib.fpText fpText5 
      Height          =   372
      Left            =   4584
      TabIndex        =   4
      Top             =   5160
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
      BackColor       =   -2147483638
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
      BorderStyle     =   1
      BorderColor     =   -2147483638
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
      AlignTextV      =   1
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
      Text            =   "Ending Date:"
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
   Begin EditLib.fpText fpText6 
      Height          =   372
      Left            =   4824
      TabIndex        =   5
      Top             =   4560
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
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
      BackColor       =   -2147483638
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
      BorderStyle     =   1
      BorderColor     =   -2147483638
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
      AlignTextV      =   1
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
      Text            =   "Start Date:"
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
   Begin EditLib.fpText fpText4 
      Height          =   372
      Left            =   3984
      TabIndex        =   3
      Top             =   3960
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
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
      BackColor       =   -2147483638
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
      BorderStyle     =   1
      BorderColor     =   -2147483638
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
      AlignTextV      =   1
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
      Text            =   "Last Employee No:"
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
   Begin EditLib.fpText fpText3 
      Height          =   372
      Left            =   3984
      TabIndex        =   2
      Top             =   3360
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
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
      BackColor       =   -2147483638
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
      BorderStyle     =   1
      BorderColor     =   -2147483638
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
      AlignTextV      =   1
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
      Text            =   "First Employee No:"
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
   Begin EditLib.fpText fpText2 
      Height          =   732
      Left            =   4224
      TabIndex        =   1
      Top             =   2160
      Width           =   5412
      _Version        =   196608
      _ExtentX        =   9546
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483648
      ForeColor       =   65535
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483643
      BorderWidth     =   3
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
      AlignTextV      =   1
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
      Text            =   " Employee Earnings History Report"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483643
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
   Begin EditLib.fpText fpText1 
      Height          =   6372
      Left            =   3144
      TabIndex        =   0
      Top             =   1440
      Width           =   7428
      _Version        =   196608
      _ExtentX        =   13102
      _ExtentY        =   11239
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   3
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   3
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483630
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
      ThreeDText      =   4
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
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483640
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   -1  'True
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   1
      BorderDropShadowColor=   -2147483634
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6612
      Left            =   3000
      Top             =   1320
      Width           =   7692
   End
End
Attribute VB_Name = "frmEmpHistRptSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   Unload frmEmpHistRptSplash
End Sub
Private Sub cmdProcess_Click()


  ReDim TempScrn(1)
  ReDim EMP2Rec(1) As EmpData2Type
  ReDim Emp1Rec(1) As EmpData1Type
  ReDim TransRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim EMPHIST(1 To 3) As EmpHistoryRptType
  ReDim DedCodes(1 To 12) As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  ReDim EmpHistRpt(1) As EmpHistFormType

  ReDim DashLine(1) As String * 132
  ReDim TotDeds(1 To 12) As Double
  ReDim TotErns(1 To 3) As Double
  ReDim ESubDeds(1 To 12) As Double
  ReDim ESubErns(1 To 3) As Double
  ReDim EmpNo(1) As String * 14
  ReDim RErnP(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim OErnP(1) As String * 11
  ReDim Ded(1) As String * 11
  ReDim Ern(1) As String * 11
  ReDim Pg(1) As String * 5
  ReDim Fill11(1) As String * 11
  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim THrs(1) As String * 11

  'REDIM OTHrs(1)             AS STRING * 11

  ReDim PHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim RErnP(1) As String * 11
  ReDim EChkDate(1) As String * 11
  ReDim EChkNo(1) As String * 11
  
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  ReDim DedCodes(1 To 12) As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType

  ReDim TFedGrs(1) As String * 11
  ReDim TStaGrs(1) As String * 11
  ReDim TSocGrs(1) As String * 11
  ReDim TMedGrs(1) As String * 11
  ReDim TRetGrs(1) As String * 11
  Dim Image2 As String, Image3 As String
  Dim UnitHandle As Integer
  Dim City As String
  Dim ErnCodeFileHandle%, DedCodeFileHandle%, x%, cnt%, LastDed%, LastErn%
  Dim DTitle$, TDed$, ETitle$, TErn$, SumHeader2$
  Dim FirstEmp&, LastEmp&, StartDate%, EndDate%
  
  Dim EmpRecSize%, TransRecLen%, LineCnt%, MaxLines%, Page%, IdxRecLen%
  Dim NumOfRecs%, EmpIdxLNameHandle%, Emp1RecLen%, EHandle1%
  Dim IdxFileSize&, Today$, FromToDate$, SumFlag%
  Dim RptTitle$, RHandle%
  Dim EmpHistoryRpt$, UsingThisOne As Boolean
  Dim THandle%, DHandle%, RecNo%
  Dim EmpHistHeader As Boolean
  Dim TTaxFring#, TFedGross#, TStaGross#, TSocGross#, TMedGross#, TRetGross#
  Dim TaxFring#, FedGross#, STAGROSS#, SocGross#, MedGross#, RETGROSS#, DAmt#
  Dim TransRecNum&, SalCnt%, HrlCnt%
  Dim FF$, SumDed$, Cnt2%, SumErn$
  
  FF$ = Chr$(12)
  Image2$ = "####.##"
  Image3$ = "########.##"

  LSet Fill11(1) = ""

  LSet DashLine(1) = String$(132, "-")
  
   OpenUnitFile UnitHandle
   Get UnitHandle, 1, UnitFileRec
   City = QPTrim$(UnitFileRec.UFEMPR)
   Close UnitHandle
     
   OpenDedCodeFile DedCodeFileHandle
   For x = 1 To 12
      Get DedCodeFileHandle, x, DedCodes(x)
   Next x
   Close DedCodeFileHandle
   
   OpenErnCodeFile ErnCodeFileHandle
   For x = 1 To 3
      Get ErnCodeFileHandle, x, ErnCodes(x)
   Next
   Close ErnCodeFileHandle
  

'*** Create the voluntary deduction description line
  DTitle$ = ""
  For cnt = 1 To 12
    TDed$ = QPTrim$(DedCodes(cnt).DCDESC1)
    If Len(TDed$) > 0 Then
      LastDed = LastDed + 1
      RSet Ded(1) = TDed$
      DTitle$ = DTitle$ + Ded(1)
    Else
      Exit For
    End If
  Next
  If LastDed < 12 Then
    DTitle$ = Space$(11 * (12 - LastDed)) + DTitle$
  End If

'*** Create the alternate earnings description line
  ETitle$ = ""
  For cnt = 1 To 3
    TErn$ = QPTrim$(ErnCodes(cnt).ERNCODE1)
    If Len(TErn$) > 0 Then
      LastErn = LastErn + 1
      RSet Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    Else
      Exit For
    End If
  Next
  If LastErn < 3 Then
    ETitle$ = Space$(11 * (3 - LastErn)) + ETitle$
  End If

  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"

  '------------------------------------------------------------------

  EmpRecSize = Len(EMP2Rec(1))

  TransRecLen = Len(TransRec(1))

  LineCnt = 0
  MaxLines = 50
  Page = 1

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  
  NumOfRecs = IdxFileSize& \ IdxRecLen

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  'If NumOfRecs = 0 Then
  '   MsgBox "No employee entries found"
  '   GoTo EndTrans
  'End If
  
  'load ThisSort with employee list in alphabetical order
  'ReDim ThisSort(EmpIdxLNameCnt)
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  Emp1RecLen = Len(Emp1Rec(1))
  
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1Rec(1)
  EmpHistRpt(1).FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1Rec(1)
  EmpHistRpt(1).LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1

  Today$ = Date$
  'EmpHistRpt(1).EndDate = Date2Num(Today$)
  'EmpHistRpt(1).StartDate = Date2Num("01-01-20" + Right$(Today$, 2))
  EmpHistRpt(1).SumOnly = "N"

  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtLastEmpNo.Text)
  StartDate = DateDiff("d", "12/31/1979", fptxtStartDate.Text)
  EndDate = DateDiff("d", "12/31/1979", fptxtEndDate.Text)
  
  FromToDate$ = "Report Date: " + QPTrim$(fptxtStartDate.Text) + " to " + QPTrim$(fptxtEndDate.Text)

  If InStr("Yy", QPTrim$(fptxtSummary.Text)) Then
    SumFlag = True
  Else
    SumFlag = False
  End If

  If LastEmp& < FirstEmp& Or EndDate < StartDate Then
    'ADD ERROR TRAP
  End If
  
  'If ExitFlag Then GoTo AltExit

'**********************************************************
'------------------------------------------------------------------
  
  
  RptTitle$ = "Employee Earnings History Report"
  
  EmpHistoryRpt = "EMPHIST.RPT"
  
  'ShowProcessingScrn RptTitle$
  RHandle = FreeFile
  Open EmpHistoryRpt For Output As RHandle

  'RPTSetupPRN 4, RHandle
  
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TransRecLen
'***
'  THistRecs = FileSize(TransHistFileName) \ TransRecLen
'***
  
  DHandle = FreeFile
  OpenEmpData2File DHandle
  
  EmpHistHeader = False

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    If Not SumFlag Then
      EmpHistHeader = False
    End If
    DAmt# = 0
    
    Get DHandle, IdxBuff(RecNo), EMP2Rec(1)

    If Val(EMP2Rec(1).EmpNo) >= FirstEmp& And Val(EMP2Rec(1).EmpNo) <= LastEmp& Then
    'if employee number is in range
      If EMP2Rec(1).LastTransRec > 0 Then         'if there are any
        TransRecNum& = EMP2Rec(1).LastTransRec
      Else
        GoTo Skip2NextEmp
      End If
      Do
        Get THandle, TransRecNum&, TransRec(1)
'***
'HistRecs = HistRecs + 1
'***
        If (TransRec(1).CheckDate >= StartDate) And (TransRec(1).CheckDate <= EndDate) Then
        'if this is in the date range
          UsingThisOne = True
          GoSub PrintAndSumEmp
          If LineCnt >= MaxLines Then
             
            Print #RHandle, FF$
            LineCnt = 0
            GoSub PrintEmpHistoryHeader
          End If
        End If 'ELSE
          If TransRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransRec(1).PrevTransRec
          Else
            If UsingThisOne Then
              GoSub PrintSubTotal
              Exit Do
            Else
              GoTo Skip2NextEmp
            End If
          End If
        'END IF
      Loop
      EMPHIST(1) = EMPHIST(2)
      ReDim ESubDeds(1 To 12) As Double
      ReDim ESubErns(1 To 3) As Double

      TTaxFring# = Round#(TTaxFring# + TaxFring#)
      TFedGross# = Round#(TFedGross# + FedGross#)
      TStaGross# = Round#(TStaGross# + STAGROSS#)
      TSocGross# = Round#(TSocGross# + SocGross#)
      TMedGross# = Round#(TMedGross# + MedGross#)
      TRetGross# = Round#(TRetGross# + RETGROSS#)

      TaxFring# = 0
      FedGross# = 0
      STAGROSS# = 0
      SocGross# = 0
      MedGross# = 0
      RETGROSS# = 0

    End If

Skip2NextEmp:
    'ShowPctComp  RecNo, NumOfRecs

  Next

  GoSub PrintGrandTotals

  'RPTSetupPRN 0, RHandle

  
  Close THandle
  Close DHandle
  Close RHandle

'***
'  PRINT "History count;   File:"; THistRecs; " Processed:"; HistRecs
'  zz$ = INPUT$(1)
'***
  
AltExit:

  'a& = SETMEM(-1)
  'PRINT SETMEM(0)
'If Not ExitFlag Then
 ViewPrint EmpHistoryRpt, RptTitle$
'End If

Exit Sub
  
  
PrintEmpHistoryHeader:
  RSet Pg(1) = Str$(Page)
  
  Print #RHandle, QPTrim$(Unit(1).UFEMPR) + Space$(86) + "Page:" + Pg(1)
  Print #RHandle, "Employee Earnings History Report" + Space$(63) + FromToDate$
  If SumFlag Then
    Print #RHandle, DashLine(1)
    LineCnt = 4
  Else
    Print #RHandle,
    LSet EmpNo(1) = QPTrim$(EMP2Rec(1).EmpNo)
    Print #RHandle, EmpNo(1) + QPTrim$(EMP2Rec(1).EMPLNAME) + ", " + QPTrim$(EMP2Rec(1).EMPFNAME)
    Print #RHandle, " Trans Date   Check No  Tax Fring   Reg Hrs      Vacat       Sick        Hol       Comp    Personal      Total   O/T Paid        EIC"
    Print #RHandle, ETitle$
    Print #RHandle, DTitle$
    Print #RHandle, DashLine(1)
    LineCnt = 9
    Page = Page + 1
  End If
  
  Return
  
PrintAndSumEmp:
  If Not EmpHistHeader Then
    EmpHistHeader = True
    GoSub PrintEmpHistoryHeader
  End If

  EMPHIST(1).RegHrs = Round#(EMPHIST(1).RegHrs + TransRec(1).RegHrsWork)
  EMPHIST(1).VACHRS = Round#(EMPHIST(1).VACHRS + TransRec(1).VacUsed)
  EMPHIST(1).SICKHRS = Round#(EMPHIST(1).SICKHRS + TransRec(1).SickUsed)
  EMPHIST(1).HOLHRS = Round#(EMPHIST(1).HOLHRS + TransRec(1).HOLHOURS)
  EMPHIST(1).COMPHRS = Round#(EMPHIST(1).COMPHRS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = Round(EMPHIST(1).TotalHrs + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = Round(EMPHIST(1).TotalHrs + TransRec(1).PerHours)
  
  'EMPHIST(1).TOTHrs = Round#(EMPHIST(1).TOTHrs + TransRec(1).OTHours)

  EMPHIST(1).PHrs = Round#(EMPHIST(1).PHrs + TransRec(1).PerHours)

  EMPHIST(1).TOTPaid = Round#(EMPHIST(1).TOTPaid + TransRec(1).OTHrsPaid)
  EMPHIST(1).TOTEIC = Round#(EMPHIST(1).TOTEIC + TransRec(1).EICAmt)
  
  EMPHIST(1).TRegWage = Round#(EMPHIST(1).TRegWage + TransRec(1).TotRegWage)
  EMPHIST(1).TOTWage = Round#(EMPHIST(1).TOTWage + TransRec(1).TotOTWage)
  
  EMPHIST(1).GPay = Round#(EMPHIST(1).GPay + TransRec(1).GrossPay)
  EMPHIST(1).SSTax = Round#(EMPHIST(1).SSTax + TransRec(1).SocTaxAmt)
  EMPHIST(1).MTax = Round#(EMPHIST(1).MTax + TransRec(1).MedTaxAmt)
  EMPHIST(1).FTax = Round#(EMPHIST(1).FTax + TransRec(1).FedTaxAmt)
  EMPHIST(1).STax = Round#(EMPHIST(1).STax + TransRec(1).StaTaxAmt)

  TaxFring# = Round#(TaxFring# + TransRec(1).TaxFring)

  FedGross# = Round#(FedGross# + TransRec(1).FedGrossPay)
  STAGROSS# = Round#(STAGROSS# + TransRec(1).StaGrossPay)
  SocGross# = Round#(SocGross# + TransRec(1).SocGrossPay)
  MedGross# = Round#(MedGross# + TransRec(1).MedGrossPay)
  RETGROSS# = Round#(RETGROSS# + TransRec(1).RetGrossPay)

'*****
  If TransRec(1).TaxFring > 0 Then
    FedGross# = Round#(FedGross# + TransRec(1).TaxFring)
    STAGROSS# = Round#(STAGROSS# + TransRec(1).TaxFring)
    SocGross# = Round#(SocGross# + TransRec(1).TaxFring)
    MedGross# = Round#(MedGross# + TransRec(1).TaxFring)
    RETGROSS# = Round#(RETGROSS# + TransRec(1).TaxFring)
  End If

  EMPHIST(1).RETTOT = Round(EMPHIST(1).RETTOT + TransRec(1).RetireAmt)
  'TOTDED# = Round(TOTDED# + TransRec(1).TotDedAmt)
  
  EMPHIST(1).TNetPay = Round#(EMPHIST(1).TNetPay + TransRec(1).NetPay)
  
  If Not SumFlag Then
    'LSET EChkDate(1) = Num2Date$(TransRec(1).PayPdEnd)           'LTRIM$(EmpRec1(1).EmpNo)
    LSet EChkDate(1) = MakeRegDate(TransRec(1).CheckDate)           'LTRIM$(EmpRec1(1).EmpNo)
    'RSET EChkNo(1) = STR$(TransRec(1).CheckNum)   'QPRTrim$(EmpRec1(1).EMPLNAME) + ", " + QPRTrim$(EmpRec1(1).EMPFNAME)
    
    RSet EChkNo(1) = Str$(TransRecNum&)
    RSet RHrs(1) = Using(Image2$, TransRec(1).RegHrsWork)
    RSet VHrs(1) = Using(Image2$, TransRec(1).VacUsed)
    RSet SHrs(1) = Using(Image2$, TransRec(1).SickUsed)
    RSet HHrs(1) = Using(Image2$, TransRec(1).HOLHOURS)
    RSet CHrs(1) = Using(Image2$, TransRec(1).CompUsed)
    RSet THrs(1) = Using(Image2$, TransRec(1).RegHrsPaid)

    'RSET OTHrs(1) = FUsing(STR$(TransRec(1).OTHours), Image2$)

    RSet PHrs(1) = Using(Image2$, TransRec(1).PerHours)

    RSet OTPaid(1) = Using(Image2$, TransRec(1).OTHrsPaid)
    RSet EICP(1) = Using(Image2$, TransRec(1).EICAmt)
    RSet Fill11(1) = Using(Image3$, TransRec(1).TaxFring)
    RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
    RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
    RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
    RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
    RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
    RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
    RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
    RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
    RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  End If

  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  SumDed$ = ""
  For Cnt2 = 1 To LastDed
    ESubDeds(Cnt2) = Round#(ESubDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    TotDeds(Cnt2) = Round#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    RSet Ded(1) = Using(Image3$, TransRec(1).DAmt(Cnt2))
    SumDed$ = SumDed$ + Ded(1)
  Next
  If LastDed < 12 Then
    SumDed$ = Space$(11 * (12 - LastDed)) + SumDed$
  End If
  
  '----------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    ESubErns(Cnt2) = Round#(ESubErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    TotErns(Cnt2) = Round#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSet Ern(1) = Using(Image3$, TransRec(1).EAmt(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If

  '-------------------------------------------------------
'  RSET EL2(1).SumEarn = SumErn$

  If Not SumFlag Then
    
    Print #RHandle, EChkDate(1) + EChkNo(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1);
    Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
    Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1);
    Print #RHandle, RetirP(1) + NetPayP(1)
    Print #RHandle, SumDed$
    Print #RHandle,
    LineCnt = LineCnt + 4
  End If

'JumpThis:
Return
  
PrintSubTotal:
  
  RSet THrs(1) = Using(Image3$, EMPHIST(1).TotalHrs)
  RSet RHrs(1) = Using(Image3$, EMPHIST(1).RegHrs)
  RSet VHrs(1) = Using(Image3$, EMPHIST(1).VACHRS)
  RSet SHrs(1) = Using(Image3$, EMPHIST(1).SICKHRS)
  RSet HHrs(1) = Using(Image3$, EMPHIST(1).HOLHRS)
  RSet CHrs(1) = Using(Image3$, EMPHIST(1).COMPHRS)

  'RSET OTHrs(1) = FUsing(STR$(EMPHIST(1).TOTHrs), Image3$)

  RSet PHrs(1) = Using(Image3$, EMPHIST(1).PHrs)
  RSet OTPaid(1) = Using(Image3$, EMPHIST(1).TOTPaid)
  RSet EICP(1) = Using(Image3$, EMPHIST(1).TOTEIC)
  
  RSet RErnP(1) = Using(Image3$, EMPHIST(1).TRegWage)
  RSet OErnP(1) = Using(Image3$, EMPHIST(1).TOTWage)
  
  RSet GPayP(1) = Using(Image3$, EMPHIST(1).GPay)
  RSet SSTaxP(1) = Using(Image3$, EMPHIST(1).SSTax)
  RSet MTaxP(1) = Using(Image3$, EMPHIST(1).MTax)
  RSet FTaxP(1) = Using(Image3$, EMPHIST(1).FTax)
  RSet STaxP(1) = Using(Image3$, EMPHIST(1).STax)
  RSet RetirP(1) = Using(Image3$, EMPHIST(1).RETTOT)
  RSet NetPayP(1) = Using(Image3$, EMPHIST(1).TNetPay)

  RSet Fill11(1) = Using(Image3$, TaxFring#)
  RSet TFedGrs(1) = Using(Image3$, FedGross#)
  RSet TStaGrs(1) = Using(Image3$, STAGROSS#)
  RSet TSocGrs(1) = Using(Image3$, SocGross#)
  RSet TMedGrs(1) = Using(Image3$, MedGross#)
  RSet TRetGrs(1) = Using(Image3$, RETGROSS#)

  '---------------------------------------------------------------
  
  EMPHIST(3).TotalHrs = Round(EMPHIST(3).TotalHrs + EMPHIST(1).TotalHrs)
  EMPHIST(3).RegHrs = Round(EMPHIST(3).RegHrs + EMPHIST(1).RegHrs)
  EMPHIST(3).VACHRS = Round(EMPHIST(3).VACHRS + EMPHIST(1).VACHRS)
  EMPHIST(3).SICKHRS = Round(EMPHIST(3).SICKHRS + EMPHIST(1).SICKHRS)
  EMPHIST(3).HOLHRS = Round(EMPHIST(3).HOLHRS + EMPHIST(1).HOLHRS)
  EMPHIST(3).COMPHRS = Round(EMPHIST(3).COMPHRS + EMPHIST(1).COMPHRS)
  'EMPHIST(3).TOTHrs = Round(EMPHIST(3).TOTHrs + EMPHIST(1).TOTHrs)
  EMPHIST(3).PHrs = Round(EMPHIST(3).PHrs + EMPHIST(1).PHrs)
  EMPHIST(3).TOTPaid = Round(EMPHIST(3).TOTPaid + EMPHIST(1).TOTPaid)
  EMPHIST(3).TOTEIC = Round(EMPHIST(3).TOTEIC + EMPHIST(1).TOTEIC)
  EMPHIST(3).TRegWage = Round(EMPHIST(3).TRegWage + EMPHIST(1).TRegWage)
  EMPHIST(3).TOTWage = Round(EMPHIST(3).TOTWage + EMPHIST(1).TOTWage)
  EMPHIST(3).GPay = Round(EMPHIST(3).GPay + EMPHIST(1).GPay)
  EMPHIST(3).SSTax = Round(EMPHIST(3).SSTax + EMPHIST(1).SSTax)
  EMPHIST(3).MTax = Round(EMPHIST(3).MTax + EMPHIST(1).MTax)
  EMPHIST(3).FTax = Round(EMPHIST(3).FTax + EMPHIST(1).FTax)
  EMPHIST(3).STax = Round(EMPHIST(3).STax + EMPHIST(1).STax)
  EMPHIST(3).RETTOT = Round(EMPHIST(3).RETTOT + EMPHIST(1).RETTOT)
  EMPHIST(3).TNetPay = Round(EMPHIST(3).TNetPay + EMPHIST(1).TNetPay)
  '---------------------------------------------------------------
  
  SumDed$ = ""
  For Cnt2 = 1 To LastDed
    RSet Ded(1) = Using(Image3$, ESubDeds(Cnt2))
    SumDed$ = SumDed$ + Ded(1)
  Next
  If LastDed < 12 Then
    SumDed$ = Space$(11 * (12 - LastDed)) + SumDed$
  End If
  '---------------------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, ESubErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  
  '--------------NEW----------------------------
  RSet Pg(1) = Str$(Page)
  If Not SumFlag Then
    Print #RHandle, DashLine(1)
  End If
  If SumFlag Then
    LSet EmpNo(1) = QPTrim$(EMP2Rec(1).EmpNo)
    Print #RHandle, EmpNo(1) + QPTrim$(EMP2Rec(1).EMPLNAME) + ", " + QPTrim$(EMP2Rec(1).EMPFNAME)
    Print #RHandle, "                        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  Else
    Print #RHandle, "Employee Totals:        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  End If

  Print #RHandle, Space$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1);
'  FPut RHandle, "Employee Totals:      " + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1)
  
  Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1);
  Print #RHandle, STaxP(1) + RetirP(1) + NetPayP(1)
  Print #RHandle,
  Print #RHandle, DTitle$
  Print #RHandle, SumDed$
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross  Ret Gross"
  Print #RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + TRetGrs(1)
  Print #RHandle,
  Print #RHandle, DashLine(1)
  If Not SumFlag Then
    
    Print #RHandle, FF$
    LineCnt = 0
  Else
    LineCnt = LineCnt + 14
  End If
Return
  
  
  '-----------------------------------------------------------------------
PrintGrandTotals:
  RSet Fill11(1) = Using(Image3$, TTaxFring#)
  RSet THrs(1) = Using(Image3$, EMPHIST(3).TotalHrs)
  RSet RHrs(1) = Using(Image3$, EMPHIST(3).RegHrs)
  RSet VHrs(1) = Using(Image3$, EMPHIST(3).VACHRS)
  RSet SHrs(1) = Using(Image3$, EMPHIST(3).SICKHRS)
  RSet HHrs(1) = Using(Image3$, EMPHIST(3).HOLHRS)
  RSet CHrs(1) = Using(Image3$, EMPHIST(3).COMPHRS)

  'RSET OTHrs(1) = FUsing(STR$(EMPHIST(3).TOTHrs), Image3$)

  RSet PHrs(1) = Using(Image3$, EMPHIST(3).PHrs)
  RSet OTPaid(1) = Using(Image3$, EMPHIST(3).TOTPaid)
  RSet EICP(1) = Using(Image3$, EMPHIST(3).TOTEIC)
  
  RSet RErnP(1) = Using(Image3$, EMPHIST(3).TRegWage)
  RSet OErnP(1) = Using(Image3$, EMPHIST(3).TOTWage)
  
  RSet GPayP(1) = Using(Image3$, EMPHIST(3).GPay)
  RSet SSTaxP(1) = Using(Image3$, EMPHIST(3).SSTax)
  RSet MTaxP(1) = Using(Image3$, EMPHIST(3).MTax)
  RSet FTaxP(1) = Using(Image3$, EMPHIST(3).FTax)
  RSet STaxP(1) = Using(Image3$, EMPHIST(3).STax)
  RSet RetirP(1) = Using(Image3$, EMPHIST(3).RETTOT)
  RSet NetPayP(1) = Using(Image3$, EMPHIST(3).TNetPay)

  RSet TFedGrs(1) = Using(Image3$, TFedGross#)
  RSet TStaGrs(1) = Using(Image3$, TStaGross#)
  RSet TSocGrs(1) = Using(Image3$, TSocGross#)
  RSet TMedGrs(1) = Using(Image3$, TMedGross#)
  RSet TRetGrs(1) = Using(Image3$, TRetGross#)

  '---------------------------------------------------------------
  SumDed$ = ""
  For Cnt2 = 1 To LastDed
    RSet Ded(1) = Using(Image3$, TotDeds(Cnt2))
    SumDed$ = SumDed$ + Ded(1)
  Next
  If LastDed < 12 Then
    SumDed$ = Space$(11 * (12 - LastDed)) + SumDed$
  End If
  '---------------------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, TotErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  
  '--------------NEW----------------------------
  If SumFlag Then
    Print #RHandle, FF$
  End If
  RSet Pg(1) = Str$(Page)
  
  Print #RHandle, City + Space$(86) + "Page:" + Pg(1)
  Print #RHandle, "Employee Earnings History Report" + Space$(63) + FromToDate$
  Print #RHandle, DashLine(1)
  Print #RHandle,
  Print #RHandle, "Report Totals:          Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  Print #RHandle, Space$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1);
  Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"

'add grand totals here
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1);
  Print #RHandle, STaxP(1) + RetirP(1) + NetPayP(1)

  Print #RHandle,
  Print #RHandle, DTitle$
  Print #RHandle, SumDed$
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross  Ret Gross"
  Print #RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + TRetGrs(1)
  
  Print #RHandle, FF$
  
  Return
  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  SetupHistReportForm
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub SetupHistReportForm()

   Dim EmpData1Handle As Integer, EmpIdxLNameHandle As Integer
   Dim EmpData1Rec As EmpData1Type
   Dim IdxRecPointer As Integer, NumOfRecs As Integer
   OpenEmpData1File EmpData1Handle
   OpenEmpIdxNNameFile EmpIdxLNameHandle
   NumOfRecs = LOF(EmpIdxLNameHandle) / 2
   
   Get #EmpIdxLNameHandle, 1, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
   fptxtFirstEmpNo.Text = Val(EmpData1Rec.EmpNo)
   
   Get #EmpIdxLNameHandle, NumOfRecs, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
'   Stop
   fptxtLastEmpNo.Text = Val(EmpData1Rec.EmpNo)
  
   Close EmpIdxLNameHandle, EmpData1Handle
   fptxtSummary = "N"
   fptxtEndDate.Text = Date$
   fptxtStartDate.Text = "01-01-" + Right$(Date$, 4)

End Sub

Private Sub fptxtEndDate_LostFocus()
   Dim DateEntry As String * 11
   DateEntry = fptxtEndDate.Text
   If Mid(DateEntry, 3, 1) <> "-" Or Mid(DateEntry, 6, 1) <> "-" Then
      MsgBox "Please enter a valid date in the Start Date field (##-##-####)."
      fptxtEndDate.SetFocus
   End If

End Sub

Private Sub fptxtStartDate_LostFocus()
   Dim DateEntry As String * 11
   DateEntry = fptxtStartDate.Text
   If Mid(DateEntry, 3, 1) <> "-" Or Mid(DateEntry, 6, 1) <> "-" Then
      MsgBox "Please enter a valid date in the Start Date field (##-##-####)."
      fptxtStartDate.SetFocus
   End If
End Sub
