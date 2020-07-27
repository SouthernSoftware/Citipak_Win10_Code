VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLAdvLetter2 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Advance Letter 2"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLAdvLetter2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleMode       =   0  'User
   ScaleWidth      =   11670.02
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8640
      Left            =   575
      TabIndex        =   12
      Top             =   24
      Width           =   8257
      _Version        =   196609
      _ExtentX        =   14564
      _ExtentY        =   15240
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483639
      Caption         =   ""
      Picture         =   "frmBLAdvLetter2.frx":08CA
      Begin EditLib.fpText fptxtTownOf 
         Height          =   300
         Left            =   2550
         TabIndex        =   0
         Top             =   165
         Width           =   3810
         _Version        =   196608
         _ExtentX        =   6720
         _ExtentY        =   529
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
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
         Text            =   "Town Of..."
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
         Height          =   300
         Left            =   2550
         TabIndex        =   1
         Top             =   480
         Width           =   3810
         _Version        =   196608
         _ExtentX        =   6720
         _ExtentY        =   529
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
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
         Text            =   "123 Main St"
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
      Begin EditLib.fpText fptxtCSZ 
         Height          =   300
         Left            =   2550
         TabIndex        =   2
         Top             =   780
         Width           =   3810
         _Version        =   196608
         _ExtentX        =   6720
         _ExtentY        =   529
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
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
         Text            =   "Town, NC 27330"
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   2355
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   480
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpMask fptxtPhone 
         Height          =   300
         Left            =   3744
         TabIndex        =   3
         Top             =   1080
         Width           =   1452
         _Version        =   196608
         _ExtentX        =   2561
         _ExtentY        =   529
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
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
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   1
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
         AllowOverflow   =   0   'False
         BestFit         =   0   'False
         ClipMode        =   0
         DataFormatEx    =   0
         Mask            =   "(###)-###-####"
         PromptChar      =   "_"
         PromptInclude   =   0   'False
         RequireFill     =   0   'False
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         AutoTab         =   0   'False
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   2610
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   2850
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   3090
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   3330
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   3570
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   3810
         Width           =   7905
         _Version        =   196608
         _ExtentX        =   13949
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         Text            =   "01234567890123456789012345678901234567890123456789012345678901234567890123456789"
         CharValidationText=   ""
         MaxLength       =   111
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
      Begin EditLib.fpText fptxtLine1 
         Height          =   375
         Index           =   7
         Left            =   4950
         TabIndex        =   11
         Top             =   5760
         Width           =   3060
         _Version        =   196608
         _ExtentX        =   5397
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
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
         MaxLength       =   50
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
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   44
         Top             =   1395
         Width           =   1410
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "The name and title of the town business license official goes here."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   285
         TabIndex        =   43
         Top             =   5445
         Width           =   3840
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   3270
         X2              =   4758
         Y1              =   5790
         Y2              =   5934
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter the body of the letter in the rows below. Keep in mind that you will have to tab between rows."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3000
         TabIndex        =   42
         Top             =   1800
         Width           =   4950
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This detachable section is designed so the addresses will appear in a windowed envelope. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2730
         TabIndex        =   40
         Top             =   6450
         Width           =   5220
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "If a customer has other charges (i.e. an outstanding balance or issuance fee)  then that information will be added here."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3510
         TabIndex        =   41
         Top             =   7650
         Width           =   4350
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Town Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1155
         TabIndex        =   32
         Top             =   7950
         Width           =   2265
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Town Of Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1155
         TabIndex        =   29
         Top             =   7710
         Width           =   2745
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State, Zip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1155
         TabIndex        =   28
         Top             =   6990
         Width           =   1695
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Total Current Privilege License Fees Due:                                  Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   2550
         TabIndex        =   16
         Top             =   7230
         Width           =   5010
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   3600
         X2              =   2112
         Y1              =   7056
         Y2              =   8112
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":08E6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   5205
         Width           =   7935
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   8160
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmBLAdvLetter2.frx":0974
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   744
         Left            =   2544
         TabIndex        =   15
         Top             =   4368
         Width           =   4560
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":0A09
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   38
         Top             =   4920
         Width           =   7935
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":0A9C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Top             =   4725
         Width           =   7935
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":0B2F
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Top             =   4530
         Width           =   7935
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":0BC2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   35
         Top             =   4350
         Width           =   7935
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State, Zip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1155
         TabIndex        =   33
         Top             =   8100
         Width           =   3945
      End
      Begin VB.Line Line8 
         X1              =   7974
         X2              =   4950
         Y1              =   5685
         Y2              =   5685
      End
      Begin VB.Line Line7 
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   8208
         Y1              =   6165
         Y2              =   6165
      End
      Begin VB.Label lblPay 
         BackColor       =   &H80000009&
         Caption         =   "Make checks payable to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1155
         TabIndex        =   31
         Top             =   7470
         Width           =   3135
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   25
         Top             =   6315
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter the town's address and phone number in these fields."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   6480
         TabIndex        =   18
         Top             =   288
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Today's date and customer data are automatically populated at runtime."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1164
         Left            =   336
         TabIndex        =   13
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1968
         Width           =   2268
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State, Zip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   2268
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8160
         Y1              =   4125
         Y2              =   4125
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAdvLetter2.frx":0C55
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   255
         TabIndex        =   19
         Top             =   4155
         Width           =   7935
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   1872
         X2              =   1776
         Y1              =   864
         Y2              =   1968
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   3114
         X2              =   2010
         Y1              =   2040
         Y2              =   2868
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   6768
         X2              =   5472
         Y1              =   1152
         Y2              =   1248
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   3840
         X2              =   1440
         Y1              =   1488
         Y2              =   864
      End
      Begin VB.Label Head 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Today's Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Index           =   0
         Left            =   3600
         TabIndex        =   17
         Top             =   1392
         Width           =   1752
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1776
         Width           =   2268
      End
      Begin VB.Label lblDetach 
         BackColor       =   &H80000009&
         Caption         =   "Detach and return this portion with your remittance. Thank you."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2835
         TabIndex        =   30
         Top             =   6165
         Width           =   3135
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   3312
         X2              =   2160
         Y1              =   6816
         Y2              =   6960
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1155
         TabIndex        =   27
         Top             =   6795
         Width           =   2265
      End
      Begin VB.Label Return 
         BackColor       =   &H80000009&
         Caption         =   "Business Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1155
         TabIndex        =   26
         Top             =   6585
         Width           =   2265
      End
      Begin VB.Label Head 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1584
         Width           =   1644
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   510
      Left            =   9435
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1590
      _Version        =   131072
      _ExtentX        =   2805
      _ExtentY        =   900
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
      ButtonDesigner  =   "frmBLAdvLetter2.frx":0CE8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   510
      Left            =   9435
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1590
      _Version        =   131072
      _ExtentX        =   2805
      _ExtentY        =   900
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
      ButtonDesigner  =   "frmBLAdvLetter2.frx":0EC6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   675
      Left            =   9435
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1590
      _Version        =   131072
      _ExtentX        =   2805
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmBLAdvLetter2.frx":10A2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   675
      Left            =   9423
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1603
      _Version        =   131072
      _ExtentX        =   2828
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmBLAdvLetter2.frx":1284
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   480
      Left            =   9300
      TabIndex        =   47
      Tag             =   $"frmBLAdvLetter2.frx":1466
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   3696
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmBLAdvLetter2.frx":1530
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   870
      Left            =   9108
      Top             =   3480
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1212
      Left            =   9504
      Top             =   1920
      Width           =   1404
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Letter #2"
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
      Height          =   684
      Left            =   9648
      TabIndex        =   45
      Top             =   2208
      Width           =   1116
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnuSample 
         Caption         =   "Print Sample"
      End
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Print Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmBLAdvLetter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  If Exist("issueappslics.dat") Then
    KillFile "issueappslics.dat"
    frmBLIssueAppsLics.Show
    DoEvents
  End If
  frmBLTownSetup.fpcmbLaserYN.SetFocus
  Unload frmBLAdvLetter2
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
    Line5.Visible = True
    Line6.Visible = True
    Line10.Visible = True
    Line11.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Line5.Visible = False
    Line6.Visible = False
    Line10.Visible = False
    Line11.Visible = False
  End If
End Sub

Private Sub cmdLast_Click()
  frmBLAdvanceLetter.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdNext_Click()
  frmBLAdvanceLtr3.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  Dim LaserRec2 As LaserLetterType2
  Dim LHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    TownRec.LaserLtr = "2"
    Put THandle, 1, TownRec
  Else
    TownRec.AppForm = 11
    TownRec.TownName = ""
    TownRec.Contact = ""
    TownRec.TownAdd1 = ""
    TownRec.TownAdd2 = ""
    TownRec.City = ""
    TownRec.State = ""
    TownRec.ZipCode = ""
    TownRec.TownPhone = ""
    TownRec.SpareSpace = ""
    TownRec.AppForm = 1
    TownRec.DLQNotice = 0
    TownRec.AppAdd1 = ""
    TownRec.AppBaseFee(1) = 0
    TownRec.AppBaseFee(2) = 0
    TownRec.AppBaseFee(3) = 0
    TownRec.AppBaseFee(4) = 0
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0
    TownRec.AppFirstDay = ""
    TownRec.AppLastDay = ""
    TownRec.AppGrsRcpts(1) = 0
    TownRec.AppGrsRcpts(2) = 0
    TownRec.AppGrsRcpts(3) = 0
    TownRec.AppGrsRcpts(4) = 0
    TownRec.AppColFee = 0
    TownRec.AppGrsPct = 0
    TownRec.AppDenom = 0
    TownRec.AppNumer = 0
    TownRec.AppState = ""
    TownRec.AppCity = ""
    TownRec.AppTownOf = ""
    TownRec.AppZip = ""
    TownRec.AppPct = 0
    TownRec.AppPayBy = 0
    TownRec.AppPhone = ""
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppDiscPct = 0
    TownRec.AppDiscMonth = ""
    TownRec.AppDiscDay = 0
    TownRec.AppPenMonth = ""
    TownRec.AppPenDay = 0
    TownRec.AppFiscMonth = ""
    TownRec.AppFiscDay = 0
    TownRec.AppMayorCouncil = ""
    TownRec.AppWholeMonth = 0
    TownRec.AppWholeDay = 0
    TownRec.AppRetailMonth = 0
    TownRec.AppRetailDay = 0
    TownRec.AppFinMonth = 0
    TownRec.AppFinDay = 0
    TownRec.AppContMonth = 0
    TownRec.AppContDay = 0
    TownRec.AppRepairMonth = 0
    TownRec.AppRepairDay = 0
    TownRec.AppStartMonth = ""
    TownRec.AppStartDay = 0
    TownRec.AppLicRetMonth = ""
    TownRec.AppLicRetDay = 0
    TownRec.AppAdoptDate = 0
    TownRec.AppCityOrd = ""
    For x = 1 To 10
      TownRec.AppYrUpDown(x) = "0"
    Next x
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = ""
    TownRec.DlqAdminTitle = ""
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = ""
    TownRec.DlqFax = ""
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = ""
    TownRec.LicNumPermYN = "No"
    TownRec.UseAmtPctYN = "Pct"
    TownRec.PENCASHACCT = 0
    TownRec.PENRECGLNUM = 0
    TownRec.PENREVGLNUM = 0
    TownRec.IssFee = 0
    TownRec.AcctMeth = ""
    TownRec.LaserLtr = "2"
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
  Close THandle
  
  OpenLaserFile2 LHandle
  
  LaserRec2.TownOf = QPTrim$(fptxtTownOf.Text)
  LaserRec2.Address = QPTrim$(fptxtAddress.Text)
  LaserRec2.CityStateZip = QPTrim$(fptxtCSZ.Text)
  LaserRec2.Phone = QPTrim$(fptxtPhone.Text)
  For x = 0 To 7
    LaserRec2.Line1(x) = fptxtLine1(x).Text
  Next x
  
  Put LHandle, 1, LaserRec2
  
  Close LHandle
    
  frmBLSucSave.Label1.Caption = "Your advance renewal letter #2 has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  Call cmdExit_Click
  frmBLTownSetup.fpcmbLaserYN.Text = "2"
  
  MainLog ("Business license advance letter #2 created and saved.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdvanceLetter", "cmdSave_Click", Erl)
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
    ClearInUse PWcnt
    Terminate
  
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Call FixFonts
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%N"
      Call cmdNext_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%L"
      Call cmdLast_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLAdvLetter2.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim LaserRec2 As LaserLetterType2
  Dim LaserHandle As Integer
  Dim City$
  Dim State$
  Dim Zip$
  
  On Error GoTo ERRORSTUFF
  
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Line2.Visible = False
  Line3.Visible = False
  Line4.Visible = False
  Line5.Visible = False
  Line6.Visible = False
  Line10.Visible = False
  Line11.Visible = False
  'this screen could not be opened unless the Town Setup
  'Laser option is either 1 or 2
  If Exist("arlaser2.dat") Then 'it's wanted but has it been saved yet?
    OpenLaserFile2 LaserHandle
    Get LaserHandle, 1, LaserRec2
    Close LaserHandle
    fptxtTownOf.Text = QPTrim$(LaserRec2.TownOf)
    Me.Return(4).Caption = QPTrim$(LaserRec2.TownOf)
    fptxtAddress.Text = QPTrim$(LaserRec2.Address)
    Me.Return(5).Caption = QPTrim$(LaserRec2.Address)
    fptxtCSZ.Text = QPTrim$(LaserRec2.CityStateZip)
    Me.Return(7).Caption = QPTrim$(LaserRec2.CityStateZip)
    fptxtPhone.Text = QPTrim$(LaserRec2.Phone)
    For x = 0 To 7
      If x = 7 Then
        fptxtLine1(x).Text = QPTrim$(LaserRec2.Line1(x))
      Else
        fptxtLine1(x).Text = LaserRec2.Line1(x)
      End If
    Next x
  Else 'laser.dat exists but #2 hasn't been saved
    fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    Me.Return(4).Caption = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    fptxtAddress.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
    Me.Return(5).Caption = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    fptxtCSZ.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text) + ", " + _
      QPTrim$(frmBLTownSetup.fptxtState.Text) + " " + QPTrim$(frmBLTownSetup.fptxtZip)
    Me.Return(7).Caption = QPTrim$(fptxtCSZ.Text)
    For x = 0 To 7
      fptxtLine1(x).Text = ""
    Next x
  End If
    
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdvLetter2", "LoadMe", Erl)
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
    ClearInUse PWcnt
    Terminate
  
End Sub

Private Sub FixFonts()
  Dim x As Integer
  
  On Error Resume Next
  Select Case ScreenW
    Case 1280
'      Shape1.Left = 9540
'      Label8.Left = 9684
'      Shape1.Top = 1920
'      Label8.Top = 2208
'      fptxtTownOf.Left = 2070
'      fptxtTownOf.FontSize = 11
'      fptxtAddress.Left = 2070
'      fptxtAddress.FontSize = 10
'      fptxtCSZ.Left = 2070
'      fptxtCSZ.FontSize = 10
'      fptxtPhone.Left = 3270
'      vaImprint1.Left = 1040
'      vaImprint1.Width = 7900
'      Me.Head(0).Left = 3100
'      Me.Line1.X2 = 7695
'      Me.Line9.X2 = 7695
'      Me.Line7.X2 = 7695
'      Me.Line8.X1 = 7468
'      Me.Line8.X2 = 4477
'      fptxtLine1(7).FontSize = 7
'      fptxtLine1(7).Left = 4444
'      For x = 0 To 6
'        fptxtLine1(x).Width = 7450
'        fptxtLine1(x).FontSize = 7
'      Next x
'      For x = 0 To 7
'        Me.Return(x).Left = 600
'      Next x
'      Label3.Left = 2400
'      Label5.Left = 3200
'      Label6.Left = 2676
'      Line6.X1 = 3500
'      cmdHelp.Left = 9500
'      cmdNext.Left = 9500
'      cmdLast.Left = 9500
'      cmdExit.Left = 9500
'      cmdSave.Left = 9500
'      Label2.FontBold = False
'      Label3.FontBold = False
'      Label4.FontBold = False
'      Label5.FontBold = False
'      Label6.FontBold = False
'      Label1.FontBold = False
'      Label7.FontBold = False
    Case 1152
'      Shape1.Left = 9540
'      Label8.Left = 9684
'      Shape1.Top = 1920
'      Label8.Top = 2208
'      fptxtTownOf.Left = 2270
'      fptxtAddress.Left = 2270
'      fptxtCSZ.Left = 2270
'      fptxtPhone.Left = 3470
'      vaImprint1.Width = 8100
'      vaImprint1.Left = 840
'      Me.Line1.X2 = 7645
'      Me.Line9.X2 = 7645
'      Me.Line7.X2 = 7645
'      Me.Line8.X1 = 7468
'      Me.Line8.X2 = 4477
'      Me.Head(0).Left = 3300
'      For x = 5 To 9
'        Me.Head(x).Left = 500
'      Next x
'      For x = 0 To 6
'        fptxtLine1(x).Width = 7400
'        fptxtLine1(x).FontSize = 7
'      Next x
'      fptxtLine1(7).FontSize = 7
'      fptxtLine1(7).Left = 4444
'      Me.Return(0).FontName = "Arial Narrow"
'      Me.Return(0).FontSize = 9
'      Me.Return(1).FontName = "Arial Narrow"
'      Me.Return(1).FontSize = 7
'      Me.Return(2).FontName = "Arial Narrow"
'      Me.Return(2).FontSize = 7
'      Me.Return(3).FontName = "Arial Narrow"
'      Me.Return(3).FontSize = 7
'      Me.Return(4).FontName = "Arial Narrow"
'      Me.Return(4).FontSize = 9
'      Me.Return(5).FontName = "Arial Narrow"
'      Me.Return(5).FontSize = 7
'      Me.Return(6).FontName = "Arial Narrow"
'      Me.Return(6).FontSize = 7
'      Me.Return(7).FontName = "Arial Narrow"
'      Me.Return(7).FontSize = 7
'      cmdHelp.Left = 9500
'      cmdNext.Left = 9500
'      cmdLast.Left = 9500
'      cmdExit.Left = 9500
'      cmdSave.Left = 9500
'      Line5.X2 = 5000
'      Line6.X1 = 3500
    Case 1024
'      Shape1.Left = 9540
'      Label8.Left = 9684
'      Shape1.Top = 1920
'      Label8.Top = 2208
'      fptxtTownOf.Left = 2270
'      fptxtAddress.Left = 2270
'      fptxtCSZ.Left = 2270
'      fptxtPhone.Left = 3470
'      vaImprint1.Width = 8200
'      vaImprint1.Left = 840
'      Me.Line1.X2 = 8000
'      Me.Line9.X2 = 8000
'      Me.Line7.X2 = 8000
'      Me.Line8.X1 = 7468
'      Me.Line8.X2 = 4477
'      Me.Head(0).Left = 3300
'      For x = 0 To 6
'        fptxtLine1(x).Width = 7800
'        fptxtLine1(x).FontSize = 7.5
'      Next x
'      fptxtLine1(7).Left = 4477
'      fptxtLine1(7).FontSize = 7.5
'      Me.Return(0).FontName = "Arial Narrow"
'      Me.Return(0).FontSize = 9
'      Me.Return(1).FontName = "Arial Narrow"
'      Me.Return(1).FontSize = 7
'      Me.Return(2).FontName = "Arial Narrow"
'      Me.Return(2).FontSize = 7
'      Me.Return(3).FontName = "Arial Narrow"
'      Me.Return(3).FontSize = 7
'      Me.Return(4).FontName = "Arial Narrow"
'      Me.Return(4).FontSize = 9
'      Me.Return(5).FontName = "Arial Narrow"
'      Me.Return(5).FontSize = 7
'      Me.Return(6).FontName = "Arial Narrow"
'      Me.Return(6).FontSize = 7
'      Me.Return(7).FontName = "Arial Narrow"
'      Me.Return(7).FontSize = 7
'      cmdNext.Left = 9500
'      cmdLast.Left = 9500
'      cmdExit.Left = 9500
'      cmdSave.Left = 9500
'      cmdHelp.Left = 9500
'      Line6.X1 = 3700
    Case 800
'      Shape1.Left = 9740
'      Label8.Left = 9884
'      Shape1.Top = 1920
'      Label8.Top = 2208
'      vaImprint1.Width = 8300
'      vaImprint1.Top = 0
'      vaImprint1.Left = 840
'      For x = 0 To 6
'        fptxtLine1(x).FontName = "Arial"
'        fptxtLine1(x).FontSize = 7
'      Next x
'      fptxtLine1(7).FontName = "Arial"
'      fptxtLine1(7).FontSize = 7
'      fptxtTownOf.Left = 2350
'      fptxtTownOf.FontSize = 8.5
'      fptxtAddress.Left = 2350
'      fptxtCSZ.Left = 2350
'      fptxtPhone.Left = 3520
'      Me.Return(0).FontName = "Arial Narrow"
'      Me.Return(0).FontSize = 9
'      Me.Return(1).FontName = "Arial Narrow"
'      Me.Return(1).FontSize = 7
'      Me.Return(2).FontName = "Arial Narrow"
'      Me.Return(2).FontSize = 7
'      Me.Return(3).FontName = "Arial Narrow"
'      Me.Return(3).FontSize = 7
'      Me.Return(4).FontName = "Arial Narrow"
'      Me.Return(4).FontSize = 9
'      Me.Return(5).FontName = "Arial Narrow"
'      Me.Return(5).FontSize = 7
'      Me.Return(6).FontName = "Arial Narrow"
'      Me.Return(6).FontSize = 7
'      Me.Return(7).FontName = "Arial Narrow"
'      Me.Return(7).FontSize = 7
'      cmdNext.Left = 9560
'      cmdLast.Left = 9560
'      cmdExit.Left = 9560
'      cmdSave.Left = 9560
'      cmdHelp.Left = 9560
    Case Else
  End Select

End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("Advance letter #2 screen printed.")
End Sub

Private Sub mnuSample_Click()
  Dim ReportFile$, RptHandle As Integer
  Dim dlm$
  
  dlm = "~"
  
  ReportFile$ = "BLRPTS\ARADVLT2.RPT"
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  '                             0                        1
  Print #RptHandle, "TOWN OF PLEASANTVILLE"; dlm; "100 Main Street"; dlm;
  '                               2                         3
  Print #RptHandle, "Pleasantville, NC 28776"; dlm; "(910)-234-5678"; dlm;
  '                       4                       5                     6            7
  Print #RptHandle, "PO Box 500"; dlm; "Pleasantville Soda Shop"; dlm; Date; dlm; "#100"; dlm;
  '                          8                        9
  Print #RptHandle, "567 Elm Street"; dlm; "Pleasantville, NC 28776"; dlm;
  '                         10
  Print #RptHandle, "Dear Business Owner:"; dlm;
  '                         11
  Print #RptHandle, "     Business License renewal fees are being processed for the coming fiscal year. The Town Of Pleasantville"; dlm;
  '                         12
  Print #RptHandle, "uses these fees to enhance the business environment for all of us. For example, this year's Small Business"; dlm;
  '                         13
  Print #RptHandle, "Expo drew over 100 small businesses to the civic center to display their goods and services. Over 5,000 people"; dlm;
  '                         14
  Print #RptHandle, "attended. Please review the fee assessment for your business as outlined below. Feel free to contact this office"; dlm;
  '                         15
  Print #RptHandle, "with any concerns regarding your fees. You can expect a formal business license invoice in the next two weeks. "; dlm;
  '                         16
  Print #RptHandle, "Thank you in advance for being a member in good stannding of our town's business community."; dlm;
  '                         17
  Print #RptHandle, "Barbara Jordan, Finance Officer"; dlm;
  '                   18               19            20
  Print #RptHandle, "10000"; dlm; "Restaurant"; dlm; 25#; dlm;
  '                   21               22            23
  Print #RptHandle, "20000"; dlm; "Catering"; dlm; 15#; dlm;
  '                   24               25            26
  Print #RptHandle, "30000"; dlm; "Ice Cream Maker"; dlm; 25#; dlm;
  '                   27               28            29
  Print #RptHandle, "40000"; dlm; "Balloon Vendor"; dlm; 15#; dlm;
  '                   30               31            32
  Print #RptHandle, "50000"; dlm; "Wine-On Premise"; dlm; 20#; dlm;
  '                   33
  Print #RptHandle, 105#; dlm;
  '                 34       35
  Print #RptHandle, 0#; dlm; 0#; dlm;
  '                 36
  Print #RptHandle, 5#

  Close

  arBLAdvLetter2.Show
  
End Sub
