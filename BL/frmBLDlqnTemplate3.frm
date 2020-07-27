VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLDlqnTemplate3 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Delinquent Notice Template #3"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLDlqnTemplate3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   9375
      TabIndex        =   11
      Tag             =   "Press the 'Cancel' button to close this screen and return to the Town Setup screen."
      Top             =   6495
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmBLDlqnTemplate3.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   9375
      TabIndex        =   12
      Tag             =   "Press 'Save' to save this delinquent notice as the default delinquent notice. All fields will be committed to memory."
      Top             =   7350
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmBLDlqnTemplate3.frx":0AA8
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8640
      Left            =   1008
      TabIndex        =   13
      Top             =   48
      Width           =   7692
      _Version        =   196609
      _ExtentX        =   13568
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
      Picture         =   "frmBLDlqnTemplate3.frx":0C84
      Begin EditLib.fpText fptxtTownOf 
         Height          =   252
         Left            =   2400
         TabIndex        =   0
         Tag             =   $"frmBLDlqnTemplate3.frx":0CA0
         Top             =   624
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
         _ExtentY        =   444
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
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   38
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
      Begin EditLib.fpText fptxtSigner 
         Height          =   252
         Left            =   816
         TabIndex        =   9
         Tag             =   "In this field enter the town official most responsible for administering business licenses."
         Top             =   7056
         Width           =   2172
         _Version        =   196608
         _ExtentX        =   3831
         _ExtentY        =   444
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
         MaxLength       =   25
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
      Begin EditLib.fpText fptxtTitle 
         Height          =   252
         Left            =   816
         TabIndex        =   10
         Tag             =   "In this field enter the title of the town official most responsible for administering business licenses."
         Top             =   7296
         Width           =   2172
         _Version        =   196608
         _ExtentX        =   3831
         _ExtentY        =   444
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
      Begin EditLib.fpMask fptxtPhone 
         Height          =   300
         Left            =   4035
         TabIndex        =   3
         Tag             =   "In this field enter the telephone number of the municipal department or office responsible for administering business licenses."
         Top             =   5325
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         AlignTextH      =   0
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
      Begin EditLib.fpText fptxtCity 
         Height          =   252
         Left            =   1968
         TabIndex        =   1
         Tag             =   "Enter the town's name in this field."
         Top             =   3024
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
         _ExtentY        =   444
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
      Begin EditLib.fpText fptxtClerk 
         Height          =   275
         Left            =   1635
         TabIndex        =   2
         Tag             =   "Enter the name of a town official familiar with the administration of business licenses."
         Top             =   5325
         Width           =   2175
         _Version        =   196608
         _ExtentX        =   3836
         _ExtentY        =   485
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
         MaxLength       =   25
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
      Begin EditLib.fpText fptxtFirstDay 
         Height          =   275
         Left            =   5370
         TabIndex        =   4
         Tag             =   "In this field enter the first day of the week the municipal department or office responsible for business licenses is open."
         Top             =   5325
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   485
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
         MaxLength       =   25
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
      Begin EditLib.fpText fptxtLastDay 
         Height          =   275
         Left            =   1440
         TabIndex        =   5
         Tag             =   "In this field enter the last day of the week the municipal department or office responsible for business licenses is open."
         Top             =   5610
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   485
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
      Begin EditLib.fpText fptxtFirstHour 
         Height          =   275
         Left            =   3120
         TabIndex        =   6
         Tag             =   "In this field enter the hour the municipal department or office responsible for business licenses is open."
         Top             =   5610
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   485
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
      Begin EditLib.fpText fptxtLastHour 
         Height          =   275
         Left            =   4605
         TabIndex        =   7
         Tag             =   "In this field enter the hour the municipal department or office responsible for business licenses closes."
         Top             =   5610
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   485
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
      Begin EditLib.fpText fptxtMayorCouncil 
         Height          =   252
         Left            =   816
         TabIndex        =   8
         Tag             =   $"frmBLDlqnTemplate3.frx":0D60
         Top             =   6336
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
         _ExtentY        =   444
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
         MaxLength       =   38
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
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "issuance of any"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   41
         Top             =   3072
         Width           =   1068
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "Business License. The application form specifies"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   3312
         TabIndex        =   40
         Top             =   3072
         Width           =   3324
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000009&
         Caption         =   "OCCUPATIONAL LICENSE(BPOL) TAX. This application is required prior to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   39
         Top             =   2832
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "a XX/XX/XXXX deadline for filing, and states that a penalty may be "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   38
         Top             =   3312
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000009&
         Caption         =   "to, an audit of business records, as permitted in Code of Virginia 58.1-3939.1."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   37
         Top             =   4752
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Today's Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   2448
         TabIndex        =   36
         Tag             =   "Today's date is supplied at runtime."
         Top             =   864
         Width           =   2268
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Business Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   35
         Top             =   1104
         Width           =   1308
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   34
         Top             =   1296
         Width           =   1404
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Business Address 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   33
         Top             =   1488
         Width           =   1404
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Business City, State  Zip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   32
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "   If there are any questions or if assistance is needed in completing the form, "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   31
         Top             =   5136
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "Cordially,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   30
         Top             =   6048
         Width           =   1404
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Dear Business Owner:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   29
         Top             =   2016
         Width           =   1692
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "please call"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   870
         TabIndex        =   28
         Top             =   5370
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   "at"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   27
         Top             =   5370
         Width           =   150
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   ","
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5325
         TabIndex        =   26
         Top             =   5370
         Width           =   105
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   "through"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   870
         TabIndex        =   25
         Top             =   5655
         Width           =   540
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         Caption         =   "from"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   24
         Top             =   5655
         Width           =   345
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4410
         TabIndex        =   23
         Top             =   5655
         Width           =   210
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000009&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5910
         TabIndex        =   22
         Top             =   5655
         Width           =   210
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000009&
         Caption         =   "Town Of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   21
         Top             =   6624
         Width           =   2364
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000009&
         Caption         =   "   According to our records, your APPLICATION FOR TOWN LICENSE(S) has not "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   20
         Top             =   2352
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000009&
         Caption         =   "yet been submitted for processing your 20XX BUSINESS, PROFESSIONAL and"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   19
         Top             =   2592
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000009&
         Caption         =   "assessed on delinquent applications. The application also states a deadline of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   18
         Top             =   3552
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000009&
         Caption         =   "Weekday, Month  XX, 20XX for payment of applicable BPOL Tax, as stated"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   17
         Top             =   3792
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000009&
         Caption         =   "in Code of Virginia 58.1-3703.1-Uniform Ordinance Provisions. To avoid further"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   16
         Top             =   4032
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "action, please complete and return the APPLICATION FOR TOWN LICENSE(S)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   15
         Top             =   4272
         Width           =   5616
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000009&
         Caption         =   "immediately. Failure to comply may result in legal action including, but not limited"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   864
         TabIndex        =   14
         Top             =   4512
         Width           =   5616
         WordWrap        =   -1  'True
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   675
      Left            =   9375
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "Press this 'Next Notice' button to close this delinquent notice screen and open up the screen for delinquent notice #1."
      Top             =   4755
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
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
      ButtonDesigner  =   "frmBLDlqnTemplate3.frx":0DF7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   690
      Left            =   9375
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "Press this 'Last Notice' button to close this delinquent notice screen and open up the screen for delinquent notice #2."
      Top             =   5610
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmBLDlqnTemplate3.frx":0FD9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   480
      Left            =   9360
      TabIndex        =   45
      Tag             =   $"frmBLDlqnTemplate3.frx":11BB
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   3270
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
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
      ButtonDesigner  =   "frmBLDlqnTemplate3.frx":1285
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   9936
      TabIndex        =   46
      Top             =   1056
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   4000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   876
      Left            =   9168
      Top             =   3060
      Width           =   2268
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   9264
      TabIndex        =   47
      Top             =   4032
      Width           =   2052
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delinquent Notice #3 Virginia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   924
      Left            =   9696
      TabIndex        =   42
      Top             =   1776
      Width           =   1308
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   972
      Left            =   9312
      Top             =   1728
      Width           =   1980
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
Attribute VB_Name = "frmBLDlqnTemplate3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim ThisTown$

Private Sub cmdExit_Click()
  Unload frmBLDlqnTemplate3
  DoEvents
  frmBLTownSetup.Show
  frmBLTownSetup.fpcmbDLQNotice.SetFocus
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    frmBLMessageBoxJr.Label1.Caption = "This delinquent notice is designed for use in the State of Virginia. State Code of Virginia 58.1-3939.1 is referenced in the body of this notice."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    frmBLMessageBox.Label1.Top = 800
    frmBLMessageBox.Label1.Height = 900
    frmBLMessageBox.Label1.Caption = "Penalty amounts, X values and Today's Date will be entered in fields on the delinquent notice printing screen."
    frmBLMessageBox.Label2.Top = 2100
    frmBLMessageBox.Label2.Height = 1500
    frmBLMessageBox.Label2.Caption = "Some of the initial discretionary values appearing on this page are supplied from the Town Setup screen. If other delinquent notice templates have been used then some of the values here may have carried over from them. PLEASE REVIEW ALL values to make sure they reflect the CURRENT situation."
    frmBLMessageBox.Show vbModal
    fptxtTownOf.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtClerk.ToolTipText = ""
    fptxtPhone.ToolTipText = ""
    fptxtFirstDay.ToolTipText = ""
    fptxtLastDay.ToolTipText = ""
    fptxtFirstHour.ToolTipText = ""
    fptxtLastHour.ToolTipText = ""
    fptxtMayorCouncil.ToolTipText = ""
    fptxtSigner.ToolTipText = ""
    fptxtTitle.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdNext.ToolTipText = ""
    cmdLast.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'    fptxtCity.ToolTipText = "Enter the town name here. All town names will appear in their entirety when this form is printed."
'    fptxtClerk.ToolTipText = "Enter a town clerk's name here."
'    fptxtPhone.ToolTipText = "Enter the town's official phone number here."
'    fptxtFirstDay.ToolTipText = "Enter the first day of the week the town office is open (Monday)."
'    fptxtLastDay.ToolTipText = "Enter the last day of the week the town's office is open (Friday)."
'    fptxtFirstHour.ToolTipText = "Enter the hour the town's office opens (8:00 A.M.)."
'    fptxtLastHour.ToolTipText = "Enter the hour the town's office closes (5:00 P.M.)."
'    fptxtMayorCouncil.ToolTipText = "Enter the town's official's titles (Mayor and/or City Council)."
'    fptxtSigner.ToolTipText = "Enter the town's finance officer's name here."
'    fptxtTitle.ToolTipText = "Enter the town's finance officer's title here."
'    cmdHelp.ToolTipText = "If Help is turned on then click to deactivate the informational balloons. If turned off then press to activate instructional balloons."
'    cmdNext.ToolTipText = "Press to move to delinquent notice #1."
'    cmdLast.ToolTipText = "Press to move to delinqueunt notice #2."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = "Press to save the data on this screen."
  End If
End Sub

Private Sub cmdLast_Click()
  frmBLDlqnTemplate2a.Show
  Unload Me
End Sub

Private Sub cmdNext_Click()
  frmBLFreeFormatDlnq.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim RecNum As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Len(QPTrim$(fptxtTownOf.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save an official town name for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtTownOf.BackColor = &H80FFFF
    fptxtTownOf.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtCity.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town name for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtCity.BackColor = &H80FFFF
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtPhone.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town's telephone number for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtPhone.BackColor = &H80FFFF
    fptxtPhone.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtSigner.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save an official's name as the signer of your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtSigner.BackColor = &H80FFFF
    fptxtSigner.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtTitle.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the title of the signer of your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtTitle.BackColor = &H80FFFF
    fptxtTitle.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtClerk.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the clerk's name of your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtClerk.BackColor = &H80FFFF
    fptxtClerk.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtFirstDay.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the first day of the week the town's office is open."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFirstDay.BackColor = &H80FFFF
    fptxtFirstDay.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtLastDay.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the last day of the week the town's office is open."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtLastDay.BackColor = &H80FFFF
    fptxtLastDay.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtLastHour.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the hour the town's office closes."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtLastDay.BackColor = &H80FFFF
    fptxtLastDay.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtFirstHour.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the hour the town's office opens."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFirstDay.BackColor = &H80FFFF
    fptxtFirstDay.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtMayorCouncil.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town's official's names."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtMayorCouncil.BackColor = &H80FFFF
    fptxtMayorCouncil.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtFirstHour.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the hour the town's office opens."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFirstDay.BackColor = &H80FFFF
    fptxtFirstDay.SetFocus
    Exit Sub
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec 'other data already saved
    TownRec.DlqAdminName = QPTrim$(fptxtSigner.Text)
    TownRec.DlqAdminTitle = QPTrim$(fptxtTitle.Text)
    TownRec.DlqCity = QPTrim$(fptxtCity.Text)
    TownRec.DlqPhone = QPTrim$(fptxtPhone.Text)
    TownRec.DlqTownName = QPTrim$(fptxtTownOf.Text)
    TownRec.DLQNotice = 3
    TownRec.DlqClerkName = QPTrim$(fptxtClerk.Text)
    TownRec.DlqFirstDay = QPTrim$(fptxtFirstDay.Text)
    TownRec.DlqLastDay = QPTrim$(fptxtLastDay.Text)
    TownRec.DlqMayorCouncil = QPTrim$(fptxtMayorCouncil.Text)
    TownRec.DlqFirstHour = QPTrim$(fptxtFirstHour.Text)
    TownRec.DlqLastHour = QPTrim$(fptxtLastHour.Text)
    Put THandle, 1, TownRec
  Else
    TownRec.TownName = ""
    TownRec.Contact = ""
    TownRec.TownAdd1 = ""
    TownRec.TownAdd2 = ""
    TownRec.City = ""
    TownRec.State = ""
    TownRec.ZipCode = ""
    TownRec.TownPhone = ""
    TownRec.SpareSpace = ""
    TownRec.AppForm = 0
    TownRec.DLQNotice = 3
    TownRec.AppAdd1 = ""
    TownRec.AppBaseFee(1) = 0
    TownRec.AppBaseFee(2) = 0
    TownRec.AppBaseFee(3) = 0
    TownRec.AppBaseFee(4) = 0
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0 '20
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
    TownRec.AppPayBy = 0
    TownRec.AppState = ""
    TownRec.AppCity = ""
    TownRec.AppTownOf = ""
    TownRec.AppZip = "" '30
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppPhone = ""
    TownRec.AppPct = 0
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
    TownRec.DlqAdd1 = "" 'not used on this form
    TownRec.DlqAdminName = QPTrim$(fptxtSigner.Text)
    TownRec.DlqAdminTitle = QPTrim$(fptxtTitle.Text)
    TownRec.DlqCity = QPTrim$(fptxtCity.Text)
    TownRec.DlqPhone = QPTrim$(fptxtPhone.Text)
    TownRec.DlqPhone2 = "" 'not used on this form
    TownRec.DlqFax = "" 'not used on this form
    TownRec.DlqState = "" 'not used on this form
    TownRec.DlqTownName = QPTrim$(fptxtTownOf.Text)
    TownRec.DlqZip = "" 'not used on this form
    TownRec.DlqFirstDay = QPTrim$(fptxtFirstDay.Text)
    TownRec.DlqLastDay = QPTrim$(fptxtLastDay.Text)
    TownRec.DlqFirstHour = QPTrim$(fptxtFirstHour.Text)
    TownRec.DlqLastHour = QPTrim$(fptxtLastHour.Text)
    TownRec.DlqClerkName = QPTrim$(fptxtClerk.Text)
    TownRec.DlqMayorCouncil = QPTrim$(fptxtMayorCouncil.Text)
    TownRec.LicNumPermYN = "No"
    TownRec.UseAmtPctYN = "Pct"
    TownRec.PENCASHACCT = 0
    TownRec.PENRECGLNUM = 0
    TownRec.PENREVGLNUM = 0
    TownRec.IssFee = 0
    TownRec.AcctMeth = ""
    TownRec.LaserLtr = "N"
    TownRec.GL2Cats = "N"
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
  
  Close THandle
  frmBLSucSave.Label1.Caption = "Your delinquent notice template #3 has been saved."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  frmBLTownSetup.fpcmbDLQNotice.Text = "3. PENALTY FORM B"
  frmBLTownSetup.fpcmdDLQ.Text = "F5 Show Dl&q Notice 3"
  
  MainLog ("Delinquent template #3 saved.")
  
  Call cmdExit_Click
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDlqnTemplate3", "cmdSave_Click", Erl)
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
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call FixFonts
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
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
      SendKeys "%P"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLDlqnTemplate1.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim RecNum As Integer
  
  On Error GoTo ERRORSTUFF
  lblBalloon.Visible = False
'  fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'  fptxtCity.ToolTipText = "Enter the town name here. All town names will appear in their entirety when this form is printed."
'  fptxtClerk.ToolTipText = "Enter a town clerk's name here."
'  fptxtPhone.ToolTipText = "Enter the town's official phone number here."
'  fptxtFirstDay.ToolTipText = "Enter the first day of the week the town office is open (Monday)."
'  fptxtLastDay.ToolTipText = "Enter the last day of the week the town's office is open (Friday)."
'  fptxtFirstHour.ToolTipText = "Enter the hour the town's office opens (8:00 A.M.)."
'  fptxtLastHour.ToolTipText = "Enter the hour the town's office closes (5:00 P.M.)."
'  fptxtMayorCouncil.ToolTipText = "Enter the town's official's titles (Mayor and/or City Council)."
'  fptxtSigner.ToolTipText = "Enter the town's finance officer's name here."
'  fptxtTitle.ToolTipText = "Enter the town's finance officer's title here."
'  cmdHelp.ToolTipText = "If Help is turned on then click to deactivate the informational balloons. If turned off then press to activate instructional balloons."
'  cmdNext.ToolTipText = "Press to move to delinquent notice #1."
'  cmdLast.ToolTipText = "Press to move to delinqueunt notice #2."
'  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'  cmdSave.ToolTipText = "Press to save the data on this screen."
'  If QPTrim$(frmBLTownSetup.fpcmbAmtPct.Text) = "Amt" Then
    Label7.Caption = "a XX/XX/XXXX deadline for filing, and states that a penalty may be "
'  Else
'    Label7.Caption = "a XX/XX/XXXX deadline for filing, and states that a penalty of XX% may be "
'  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    
    If Len(QPTrim$(TownRec.DlqTownName)) = 0 Then
      If Len(QPTrim$(TownRec.TownName)) > 0 Then
        fptxtTownOf.Text = QPTrim$(TownRec.TownName)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtTownName.Text)) > 0 Then
        fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
      End If
    Else
      fptxtTownOf.Text = QPTrim$(TownRec.DlqTownName)
    End If
    
    If Len(QPTrim$(TownRec.DlqCity)) = 0 Then
      If Len(QPTrim$(TownRec.City)) > 0 Then
        fptxtCity.Text = QPTrim$(TownRec.City)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtCity.Text)) > 0 Then
        fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
      End If
    Else
      fptxtCity.Text = QPTrim$(TownRec.DlqCity)
    End If
    
    If Len(QPTrim$(TownRec.DlqPhone)) = 0 Then
      If Len(QPTrim$(TownRec.TownPhone)) > 0 Then
        fptxtPhone.Text = QPTrim$(TownRec.TownPhone)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtPhone.Text)) > 0 Then
        fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
      End If
    Else
      fptxtPhone.Text = QPTrim$(TownRec.DlqPhone)
    End If
    
    If Len(QPTrim$(TownRec.DlqClerkName)) = 0 Then
      If Len(QPTrim$(TownRec.Contact)) > 0 Then
        fptxtClerk.Text = QPTrim$(TownRec.Contact)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtContact.Text)) > 0 Then
        fptxtClerk.Text = QPTrim$(frmBLTownSetup.fptxtContact.Text)
      End If
    Else
      fptxtClerk.Text = QPTrim$(TownRec.DlqClerkName)
    End If
    
    If Len(QPTrim$(TownRec.DlqFirstDay)) = 0 Then
      fptxtFirstDay.Text = "Monday"
    Else
      fptxtFirstDay.Text = QPTrim$(TownRec.DlqFirstDay)
    End If
    
    If Len(QPTrim$(TownRec.DlqLastDay)) = 0 Then
      fptxtLastDay.Text = "Friday"
    Else
      fptxtLastDay.Text = QPTrim$(TownRec.DlqLastDay)
    End If
    
    fptxtFirstHour.Text = QPTrim$(TownRec.DlqFirstHour)
    If QPTrim$(TownRec.DlqLastHour) = "" Then
      fptxtLastHour.Text = "5:00 P.M."
    Else
      fptxtLastHour.Text = QPTrim$(TownRec.DlqLastHour)
    End If
    If QPTrim$(TownRec.DlqFirstHour) = "" Then
      fptxtFirstHour.Text = "8:00 A.M."
    Else
      fptxtFirstHour.Text = QPTrim$(TownRec.DlqFirstHour)
    End If
    fptxtSigner.Text = QPTrim$(TownRec.DlqAdminName)
    fptxtTitle.Text = QPTrim$(TownRec.DlqAdminTitle)
    fptxtClerk.Text = QPTrim$(TownRec.DlqClerkName)
    fptxtMayorCouncil.Text = QPTrim$(TownRec.DlqMayorCouncil)
    
  Else
    If QPTrim$(frmBLTownSetup.fptxtTownName.Text) = "" Then
      fptxtTownOf.Text = "Town of YourTown"
    Else
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtCity.Text) = "" Then
      fptxtCity.Text = "Your City"
    Else
      fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtPhone.Text) = "(" Then
      fptxtPhone.Text = "(555)-555-5555"
    Else
      fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtContact.Text) = "" Then
      fptxtClerk.Text = "Clerk's Name"
    Else
      fptxtClerk.Text = QPTrim$(frmBLTownSetup.fptxtContact.Text)
    End If
    
    fptxtFirstDay.Text = "Monday"
    fptxtLastDay.Text = "Friday"
    fptxtFirstHour.Text = "8:00 A.M."
    fptxtLastHour.Text = "5:00 P.M."
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDlqnTemplate3", "LoadMe", Erl)
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

Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCity.BackColor = -2147483643
End Sub

Private Sub fptxtFirstDay_Change()
  If QPTrim$(fptxtFirstDay.Text) = "" Then
    fptxtFirstDay.Text = "Monday"
  End If
End Sub

Private Sub fptxtFirstHour_Change()
  If QPTrim$(fptxtFirstHour.Text) = "" Then
    fptxtFirstHour.Text = "8:00 A.M."
  End If
End Sub

Private Sub fptxtLastDay_Change()
  If QPTrim$(fptxtLastDay.Text) = "" Then
    fptxtLastDay.Text = "Friday"
  End If
End Sub

Private Sub fptxtLastHour_Change()
  If QPTrim$(fptxtLastHour.Text) = "" Then
    fptxtLastHour.Text = "5:00 P.M."
  End If
End Sub

Private Sub fptxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtPhone.BackColor = -2147483643
End Sub

Private Sub fptxtTownName2_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtSigner.BackColor = -2147483643
End Sub

Private Sub fptxtClerk_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtClerk.BackColor = -2147483643
End Sub

Private Sub fptxtFirstDay_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtFirstDay.BackColor = -2147483643
End Sub

Private Sub fptxtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTitle.BackColor = -2147483643
End Sub

Private Sub fptxtTownOf_Change()
  If QPTrim$(fptxtTownOf.Text) = "" Then
    Label20.Caption = ThisTown
  Else
    Label20.Caption = QPTrim$(fptxtTownOf.Text)
  End If
End Sub

Private Sub fptxtTownOf_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownOf.BackColor = -2147483643
End Sub

Private Sub fptxtLastDay_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtLastDay.BackColor = -2147483643
End Sub

Private Sub fptxtLastHour_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtLastHour.BackColor = -2147483643
End Sub

Private Sub fptxtFirstHour_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtFirstHour.BackColor = -2147483643
End Sub

Private Sub fptxtMayorCouncil_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtMayorCouncil.BackColor = -2147483643
End Sub

Private Sub fptxtSigner_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtSigner.BackColor = -2147483643
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  Me.PrintForm
  MainLog ("Delinquent Notice # 3:  screen printed.")
End Sub

Private Sub FixFonts()
  Dim x As Integer
  
  On Error Resume Next
  Select Case ScreenW
    Case 1280
''      Label2.Height = 3200
'      vaImprint1.Width = 8100
'      vaImprint1.Left = 840
'      cmdHelp.Left = 9540
'      Shape1.Width = 1424
'      Shape1.Left = 9540
''      Label9.Left = 9694
'      Shape1.Top = 1920
''      Label9.Top = 2208
'      cmdHelp.Left = 9540
'      cmdNext.Left = 9500
'      cmdExit.Left = 9500
'      cmdSave.Left = 9500
''      Head(11).Left = 4500
'    Case 1152
'      vaImprint1.Width = 8100
'      vaImprint1.Left = 840
'      cmdHelp.Left = 9540
'      Shape1.Width = 1424
'      Shape1.Left = 9560
''      Label9.Left = 9694
'      Shape1.Top = 1920
''      Label9.Top = 2208
'      cmdNext.Left = 9500
'      cmdExit.Left = 9500
'      cmdSave.Left = 9500
''      Head(11).Left = 4100
    Case 1024
      Shape1.Width = 1424
      Shape1.Left = 9560
'      Label9.Left = 9694
      Shape1.Top = 1920
'      Label9.Top = 2208
      vaImprint1.Width = 7212
      vaImprint1.Left = 1728 '840
      cmdNext.Left = 9520
      cmdExit.Left = 9520
      cmdSave.Left = 9520
      cmdHelp.Left = 9560
'      Head(11).Left = 4500
    Case 800
'      vaImprint1.Top = 1
'      Shape1.Left = 9740
''      Label9.Left = 9884
'      Shape1.Top = 1920
''      Label9.Top = 2208
''      Label2.FontSize = 8
''      Label2.Height = 2200
''      Label3.FontSize = 8
''      Label5.FontSize = 8
''      Label6.FontSize = 8
''      Label6.Height = 900
''      Label7.FontSize = 8
''      Label8.FontSize = 8
''      Line2.Y1 = 2000
'      vaImprint1.Width = 8300
'      vaImprint1.Top = 0
'      vaImprint1.Left = 840
'      cmdNext.Left = 9560
'      cmdExit.Left = 9560
'      cmdSave.Left = 9560
'      cmdHelp.Left = 9600
    Case Else
  End Select

End Sub

