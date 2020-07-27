VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEnterEditPass 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Maintenance"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   12195
   Icon            =   "frmEnterEditPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
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
      Height          =   444
      Left            =   6024
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7584
      Width           =   1332
   End
   Begin EditLib.fpText fptxtConfirm 
      Height          =   396
      Left            =   7704
      TabIndex        =   3
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   1824
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   698
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
      AutoBeep        =   -1  'True
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   2
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
      CharValidationText=   "~"" `!@#$%^&*()_+-={}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin EditLib.fpText fptxtWord 
      Height          =   372
      Left            =   3792
      TabIndex        =   2
      ToolTipText     =   "No Spaces or Special Characters Allowed."
      Top             =   1848
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      AutoCase        =   1
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   2
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
      CharValidationText=   "~"" `!@#$%^&*()-_=+{}|[]\:"";'<>?,./"""
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   "*"
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
   Begin EditLib.fpLongInteger fpControlNum 
      Height          =   372
      Left            =   3816
      TabIndex        =   0
      Top             =   1320
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "1"
      MaxValue        =   "9999"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   12632256
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
   Begin EditLib.fpText fptxtUserName 
      Height          =   372
      Left            =   6384
      TabIndex        =   1
      Top             =   1296
      Width           =   3396
      _Version        =   196608
      _ExtentX        =   5990
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
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   15
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
      Height          =   444
      Left            =   7824
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7584
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
      Height          =   444
      Left            =   9624
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7584
      Width           =   1332
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   0
      Left            =   4992
      TabIndex        =   5
      Top             =   3120
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   2
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   1
      Left            =   4992
      TabIndex        =   8
      Top             =   3528
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   2
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   2
      Left            =   4992
      TabIndex        =   11
      Top             =   3936
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   2
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   3
      Left            =   4992
      TabIndex        =   14
      Top             =   4356
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   2
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   4
      Left            =   4992
      TabIndex        =   17
      Top             =   4764
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   5
      Left            =   4992
      TabIndex        =   20
      Top             =   5172
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   6
      Left            =   4992
      TabIndex        =   23
      Top             =   5580
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   7
      Left            =   4992
      TabIndex        =   26
      Top             =   5988
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   8
      Left            =   4992
      TabIndex        =   29
      Top             =   6396
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpFullAccess 
      Height          =   300
      Index           =   9
      Left            =   4992
      TabIndex        =   32
      Top             =   6792
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   " "
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   0
      Left            =   6552
      TabIndex        =   6
      Top             =   3120
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   1
      Left            =   6552
      TabIndex        =   9
      Top             =   3528
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   2
      Left            =   6552
      TabIndex        =   12
      Top             =   3936
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   3
      Left            =   6552
      TabIndex        =   15
      Top             =   4356
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   4
      Left            =   6552
      TabIndex        =   18
      Top             =   4764
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   5
      Left            =   6552
      TabIndex        =   21
      Top             =   5160
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   6
      Left            =   6552
      TabIndex        =   24
      Top             =   5580
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   7
      Left            =   6552
      TabIndex        =   27
      Top             =   5988
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   8
      Left            =   6552
      TabIndex        =   30
      Top             =   6396
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpReports 
      Height          =   300
      Index           =   9
      Left            =   6552
      TabIndex        =   33
      Top             =   6792
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   0
      Left            =   8112
      TabIndex        =   7
      Top             =   3120
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   1
      Left            =   8112
      TabIndex        =   10
      Top             =   3528
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   2
      Left            =   8112
      TabIndex        =   13
      Top             =   3936
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   3
      Left            =   8112
      TabIndex        =   16
      Top             =   4356
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   4
      Left            =   8112
      TabIndex        =   19
      Top             =   4764
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   5
      Left            =   8112
      TabIndex        =   22
      Top             =   5172
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   6
      Left            =   8112
      TabIndex        =   25
      Top             =   5580
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   0   'False
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   7
      Left            =   8112
      TabIndex        =   28
      Top             =   5988
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   8
      Left            =   8112
      TabIndex        =   31
      Top             =   6396
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpPayments 
      Height          =   300
      Index           =   9
      Left            =   8112
      TabIndex        =   34
      Top             =   6792
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   56
      Top             =   8625
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   450
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
            TextSave        =   "8:04 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/15/2018"
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
   Begin EditLib.fpBoolean fpAdmin 
      Height          =   252
      Left            =   3792
      TabIndex        =   4
      Top             =   2616
      Visible         =   0   'False
      Width           =   228
      _Version        =   196608
      _ExtentX        =   402
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   2
      ToggleFalse     =   ""
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   ""
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   1
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpAdj 
      Height          =   276
      Left            =   9720
      TabIndex        =   61
      Top             =   6408
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
      _ExtentY        =   487
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "No"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   1
      ToggleTrue      =   "Yy"
      TextTrue        =   "Yes"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "No"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Adj.)-"
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
      Height          =   276
      Left            =   8736
      TabIndex        =   62
      Top             =   6408
      Width           =   1236
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Year End Close)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   444
      Left            =   8808
      TabIndex        =   60
      Top             =   3888
      Width           =   852
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PO's"
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
      Left            =   8880
      TabIndex        =   59
      Top             =   3552
      Width           =   588
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*Administrator Has Full Access To All Modules.  Only the Administrator May Access Password Entry/Edit. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   708
      Left            =   1536
      TabIndex        =   58
      Top             =   7440
      Width           =   3948
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      Height          =   468
      Left            =   2112
      Top             =   2496
      Width           =   2076
   End
   Begin VB.Label LabelAdmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*Administrator"
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
      Height          =   348
      Left            =   2208
      TabIndex        =   57
      Top             =   2592
      Visible         =   0   'False
      Width           =   1476
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Left            =   5904
      TabIndex        =   55
      Top             =   1920
      Width           =   1644
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Height          =   348
      Left            =   2568
      TabIndex        =   54
      Top             =   1920
      Width           =   1068
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Height          =   324
      Left            =   5232
      TabIndex        =   53
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business License"
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
      Height          =   324
      Left            =   2640
      TabIndex        =   52
      Top             =   3144
      Width           =   1668
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Full Access"
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
      Left            =   4632
      TabIndex        =   51
      Top             =   2784
      Width           =   1332
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reports "
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
      Left            =   6072
      TabIndex        =   50
      Top             =   2784
      Width           =   1620
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payments/Other "
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
      Left            =   7848
      TabIndex        =   49
      Top             =   2784
      Width           =   1500
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Payable"
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
      Height          =   324
      Left            =   2496
      TabIndex        =   48
      Top             =   3552
      Width           =   1812
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "General Ledger"
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
      Height          =   324
      Left            =   2592
      TabIndex        =   47
      Top             =   3960
      Width           =   1716
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll"
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
      Height          =   324
      Left            =   3312
      TabIndex        =   46
      Top             =   4356
      Width           =   996
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Assets"
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
      Height          =   324
      Left            =   2544
      TabIndex        =   45
      Top             =   4764
      Width           =   1764
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Property Taxes"
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
      Height          =   324
      Left            =   2400
      TabIndex        =   44
      Top             =   5184
      Width           =   1908
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Control"
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
      Height          =   324
      Left            =   2688
      TabIndex        =   43
      Top             =   5580
      Width           =   1620
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Utility Billing"
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
      Height          =   324
      Left            =   2400
      TabIndex        =   42
      Top             =   6384
      Width           =   1908
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Decals"
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
      Height          =   324
      Left            =   2496
      TabIndex        =   41
      Top             =   6792
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   4644
      Left            =   2112
      Top             =   2496
      Width           =   8316
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   4512
      X2              =   4512
      Y1              =   3048
      Y2              =   7128
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   4512
      X2              =   10392
      Y1              =   3048
      Y2              =   3048
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      X1              =   6048
      X2              =   6048
      Y1              =   3048
      Y2              =   7104
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   7608
      X2              =   7608
      Y1              =   3048
      Y2              =   7128
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10392
      Y1              =   6312
      Y2              =   6312
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   5904
      Y2              =   5904
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   5496
      Y2              =   5496
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10392
      Y1              =   5088
      Y2              =   5088
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   4272
      Y2              =   4272
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10392
      Y1              =   3864
      Y2              =   3864
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   2112
      X2              =   10416
      Y1              =   3456
      Y2              =   3456
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Management"
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
      Height          =   324
      Left            =   2544
      TabIndex        =   40
      Top             =   5976
      Width           =   1764
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Control Number"
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
      Left            =   2160
      TabIndex        =   39
      Top             =   1392
      Width           =   1524
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password Maintenance"
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
      Left            =   3312
      TabIndex        =   38
      Top             =   672
      Width           =   5580
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   564
      Left            =   2580
      Top             =   576
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   2592
      Top             =   456
      Width           =   7020
   End
End
Attribute VB_Name = "frmEnterEditPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Public Recnum As Integer
Dim Citipass As CitiPassType
Dim NumPassRecs As Integer
Dim PassRecLen As Integer

Private Sub cmdDelete_Click()
  Dim cnt As Integer
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
 
  If Recnum > 0 Then
    'OpenCitiPassFile CitiPassFile, NumPassRecs
    Get CPAdminhand, Recnum, Citipass
    If Citipass.Administ Then
      MsgBox "You May Not Delete The Administrator.", vbOKOnly, "Delete Canceled."
      'Close CitiPassFile
      'If Exist("CitiPass.dat") Then SetAttr ("CitiPass.dat"), vbReadOnly
    Else
      MainLog "Delete PW " + QPTrim(Citipass.username)
      Citipass.DelFlag = True
      Put CPAdminhand, Recnum, Citipass
      'Close CitiPassFile
      'If Exist("CitiPass.dat") Then SetAttr ("CitiPass.dat"), vbReadOnly
      cmdExit_Click
    End If
  Else
    clearscrn
  End If
End Sub
Private Sub clearscrn()
  Dim cnt As Integer
  fpControlNum = ""
  fptxtUserName = ""
  fptxtWord = ""
  fptxtConfirm = ""
  For cnt = fpFullAccess.LBound To fpFullAccess.UBound
    fpFullAccess(cnt).Value = ValueTrue
  Next
  For cnt = fpReports.LBound To fpReports.UBound
    fpReports(cnt).Value = ValueTrue
  Next
  For cnt = fpPayments.LBound To fpPayments.UBound
    fpPayments(cnt).Value = ValueTrue
  Next
  Recnum = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
'    Case vbKeyF2:
'      cmdNew_Click
'      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub cmdSave_Click()
  Dim cnt As Integer
  Dim Pz As String, Z As String
  If oktosave = True Then
'  If CPAdminhand = 0 Then
'    OpenCitiPassFile CitiPassFile, NumPassRecs
'  End If
    If fpAdmin Then
      Citipass.Administ = True
      For cnt = fpFullAccess.LBound To fpFullAccess.UBound
        Citipass.Module(cnt + 1).FullAccess = True
      Next
      For cnt = fpReports.LBound To fpReports.UBound
        Citipass.Module(cnt + 1).ReportsOnly = True
      Next
      For cnt = fpPayments.LBound To fpPayments.UBound
        Citipass.Module(cnt + 1).PaymentAccess = True
      Next
      Citipass.Module(9).Adjustments = True
    Else
      Citipass.Administ = False
      For cnt = fpFullAccess.LBound To fpFullAccess.UBound
        Citipass.Module(cnt + 1).FullAccess = fpFullAccess(cnt).Value And ValueTrue
      Next
      For cnt = fpReports.LBound To fpReports.UBound
        Citipass.Module(cnt + 1).ReportsOnly = fpReports(cnt).Value And ValueTrue
      Next
      For cnt = fpPayments.LBound To fpPayments.UBound
        Citipass.Module(cnt + 1).PaymentAccess = fpPayments(cnt).Value And ValueTrue
      Next
      Citipass.Module(9).Adjustments = fpAdj.Value And ValueTrue
    End If
  Pz$ = ""
  Z$ = QPTrim(fptxtWord)
  For cnt = 1 To Len(Z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
  Next
  Citipass.PassNum = fpControlNum
  Citipass.username = fptxtUserName
  Citipass.PassWord = Pz$
  Citipass.DelFlag = False
  Citipass.Flag2 = 0
  Citipass.FlagMod = 0
  If Recnum = 0 Then
    Recnum = NumPassRecs + 1
  End If
  Put CPAdminhand, Recnum, Citipass
  MainLog "PW Added " + fpControlNum + "," + fptxtUserName
  'Close CitiPassFile
  'SetAttr ("CitiPass.dat"), vbReadOnly
  cmdExit_Click
  End If
End Sub
Private Sub cmdExit_Click()
  'Need Prompt to save info here with check*********
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  
  If NumPassRecs > 0 Then
    frmUserSelect.fpcboUsers.Clear
    frmUserSelect.fpcboUsers.Action = ActionClear
    frmUserSelect.FillUsers frmUserSelect.fpcboUsers
    frmUserSelect.fpcboUsers.ListIndex = -1
    frmUserSelect.fpcboUsers.Action = ActionClearSearchBuffer
    frmUserSelect.fpcboUsers.Enabled = True
    frmUserSelect.cmdEdit.Enabled = True
    frmUserSelect.Label3.Visible = False
  End If

  Load frmUserSelect
  frmUserSelect.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

If LevelPass = 1 Then
  frmEnterEditPass.LabelAdmin.Visible = True
  frmEnterEditPass.fpAdmin.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    Cancel = True
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Public Function Rec2Form(TempRec)
  Dim cnt As Integer
  Dim Pz As String, Z As String
  Recnum = TempRec
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  Get CPAdminhand, TempRec, Citipass
    If Citipass.Administ Then
      fpAdmin.Value = Citipass.Administ And ValueTrue
      LabelAdmin.Visible = True
      For cnt = fpFullAccess.LBound To fpFullAccess.UBound
        fpFullAccess(cnt).Value = ValueTrue
        fpFullAccess(cnt).Enabled = False
      Next
      For cnt = fpReports.LBound To fpReports.UBound
        fpReports(cnt).Value = ValueTrue
        fpReports(cnt).Enabled = False
      Next
      For cnt = fpPayments.LBound To fpPayments.UBound
        fpPayments(cnt).Value = ValueTrue
        fpPayments(cnt).Enabled = False
      Next
      fpAdj.Value = ValueTrue
      fpAdj.Enabled = False
    Else
      
      For cnt = fpFullAccess.LBound To fpFullAccess.UBound
        fpFullAccess(cnt).Value = Citipass.Module(cnt + 1).FullAccess And ValueTrue
      Next
      For cnt = fpReports.LBound To fpReports.UBound
        fpReports(cnt).Value = Citipass.Module(cnt + 1).ReportsOnly And ValueTrue
      Next
      For cnt = fpPayments.LBound To fpPayments.UBound
        fpPayments(cnt).Value = Citipass.Module(cnt + 1).PaymentAccess And ValueTrue
      Next
      fpAdj.Value = Citipass.Module(9).Adjustments And ValueTrue
      fpAdj.Enabled = True
    End If
  Pz$ = ""
  Z$ = QPTrim(Citipass.PassWord)
  For cnt = 1 To Len(Z$)
    Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
  Next
  fpControlNum = Citipass.PassNum
  fptxtUserName = Citipass.username
  fptxtWord = Pz$
  fptxtConfirm = Pz$
  'Close CitiPassFile
End Function
Private Function oktosave()
If fpControlNum > 0 Then
  If QPTrim(fptxtWord) = QPTrim(fptxtConfirm) Then
    If Chk4dup(1) = False Then
      If Chk4dup(2) = False Then
        If Chk4dup(3) = False Then
          If Chk4dup(4) = False Then
            oktosave = True
          End If
        End If
      End If
    End If
  End If
Else
  MsgBox "Operator Number 0 NOT Allowed.", vbOKOnly, "Invalid Number"
End If
End Function
Private Function Chk4dup(x As Integer)
  Dim Tellit As String
  Dim Pz As String, Z As String, cntr As Integer, found As Boolean
  Dim cnt As Integer
  Tellit = ""
  found = False
  'OpenCitiPassFile CitiPassFile, NumPassRecs
  PassRecLen = Len(Citipass)
  NumPassRecs = LOF(CPAdminhand) \ PassRecLen
  If NumPassRecs > 0 Then
  Select Case x:
    Case 1:          'control number
      For cntr = 1 To NumPassRecs
        Get CPAdminhand, cntr, Citipass
        If Not Citipass.DelFlag Then
          If fpControlNum = Citipass.PassNum Then
            If Not Recnum = cntr Then
              found = True
              Tellit = "Duplicate Control Number, Try Another."
              Exit For
            End If
          End If
        End If
      Next
    Case 2:
      For cntr = 1 To NumPassRecs
        Get CPAdminhand, cntr, Citipass
        If Not Citipass.DelFlag Then
          If QPTrim(fptxtUserName) = QPTrim(Citipass.username) Then
            If Not Recnum = cntr Then
              found = True
              Tellit = "Duplicate UserName, Must Be Unique."
              Exit For
            End If
          End If
        End If
      Next
    Case 3:
      For cntr = 1 To NumPassRecs
        Get CPAdminhand, cntr, Citipass
        If Not Citipass.DelFlag Then
          Pz$ = ""
          Z$ = QPTrim(Citipass.PassWord)
          For cnt = 1 To Len(Z$)
            Pz$ = Pz$ + Chr$(Asc(Mid$(Z$, cnt, 1)) Xor 127)
          Next
          If QPTrim(fptxtWord) = Pz$ Then
            If Not Recnum = cntr Then
              found = True
              Tellit = "Duplicate Password, Must Be Unique."
              Exit For
            End If
          End If
        End If
      Next
    Case 4:
      For cntr = 1 To NumPassRecs
        Get CPAdminhand, cntr, Citipass
        If Not Citipass.DelFlag Then
          If Citipass.Administ = True And fpAdmin.Value = ValueTrue Then
            If Not Recnum = cntr Then
              found = True
              Tellit = "Only One Administrator Allowed, Must Be Unique."
              Exit For
            End If
          End If
        End If
      Next

    Case Else:
  End Select
  'Close CitiPassFile
  If found = True Then
    MsgBox Tellit$, vbOKOnly, "Invalid Entry"
    If x = 1 Then fpControlNum.SetFocus
    If x = 2 Then fptxtUserName.SetFocus
    If x = 3 Then fptxtWord.SetFocus
  End If
  Chk4dup = found
  End If
End Function

Private Sub fpAdj_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdSave.SetFocus
  End If
End Sub

Private Sub fpAdmin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fpControlNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtUserName.SetFocus
  End If
End Sub

Private Sub fpControlNum_LostFocus()
  Chk4dup 1
End Sub

Private Sub fptxtConfirm_LostFocus()
  If QPTrim(fptxtConfirm) <> QPTrim(fptxtWord) Then
    MsgBox "Confirmation of Password does NOT Match, Please Retry.", vbOKOnly, "Invalid Entry"
    fptxtWord.SetFocus
  End If
End Sub

Private Sub fptxtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtWord.SetFocus
  End If
End Sub

Private Sub fptxtUserName_LostFocus()
  Chk4dup 2
End Sub

Private Sub fptxtWord_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtConfirm.SetFocus
  End If
End Sub
Private Sub fptxtWord_LostFocus()
  Chk4dup 3
End Sub

Private Sub fptxtConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fpFullAccess_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpReports(Index).SetFocus
  End If
End Sub
Private Sub fpReports_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpPayments(Index).SetFocus
  End If
End Sub
Private Sub fpPayments_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Index = 5 Then
      If fpFullAccess(Index + 2).Enabled Then
        fpFullAccess(Index + 2).SetFocus
      Else
        cmdSave.SetFocus
      End If
    ElseIf Index = 8 Then
      If fpAdj.Enabled = True Then
        fpAdj.SetFocus
      End If
    ElseIf Index < fpPayments.UBound Then
      If fpFullAccess(Index + 1).Enabled Then
        fpFullAccess(Index + 1).SetFocus
      Else
        cmdSave.SetFocus
      End If
    End If
  End If
End Sub


