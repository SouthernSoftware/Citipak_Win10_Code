VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmBLDlqnTemplate1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Delinquent Notice Template #1"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLDlqnTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   480
      Left            =   9360
      TabIndex        =   32
      Tag             =   $"frmBLDlqnTemplate.frx":08CA
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
      ButtonDesigner  =   "frmBLDlqnTemplate.frx":0994
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   9375
      TabIndex        =   8
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
      ButtonDesigner  =   "frmBLDlqnTemplate.frx":0B77
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   9375
      TabIndex        =   9
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
      ButtonDesigner  =   "frmBLDlqnTemplate.frx":0D55
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8640
      Left            =   1008
      TabIndex        =   10
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
      Picture         =   "frmBLDlqnTemplate.frx":0F31
      Begin EditLib.fpText fptxtTownOf 
         Height          =   275
         Left            =   2256
         TabIndex        =   0
         Tag             =   "Enter the official town name in this field. A good example of the town name for this field would be 'Town of Paradise'. "
         Top             =   820
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
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
      Begin EditLib.fpText fptxtAdd 
         Height          =   275
         Left            =   2256
         TabIndex        =   1
         Tag             =   "Enter the town's mailing address in this field."
         Top             =   1084
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
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
      Begin EditLib.fpText fptxtCity 
         Height          =   275
         Left            =   2256
         TabIndex        =   2
         Tag             =   "Enter the town's mailing name in this field."
         Top             =   1344
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
      Begin EditLib.fpText fptxtSigner 
         Height          =   252
         Left            =   3744
         TabIndex        =   6
         Tag             =   "In this field enter the name of the town official responsible for sending this notice."
         Top             =   7584
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
         Left            =   3744
         TabIndex        =   7
         Tag             =   "In this field enter the official title of the person who signed this notice."
         Top             =   7824
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
      Begin EditLib.fpText fptxtState 
         Height          =   275
         Left            =   3504
         TabIndex        =   3
         Tag             =   "Enter the town's state in this field (NC = North Carolina)"
         Top             =   1344
         Width           =   300
         _Version        =   196608
         _ExtentX        =   529
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
      Begin EditLib.fpMask fptxtPhone 
         Height          =   300
         Left            =   2730
         TabIndex        =   5
         Tag             =   "Enter the telephone number for the municipal department responsible for administering business licenses."
         Top             =   1635
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
      Begin EditLib.fpMask fptxtZip 
         Height          =   275
         Left            =   3792
         TabIndex        =   4
         Tag             =   "Enter the town's postal code in this field."
         Top             =   1344
         Width           =   876
         _Version        =   196608
         _ExtentX        =   1545
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
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         Caption         =   "Account #: XXX"
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
         Left            =   4464
         TabIndex        =   28
         Top             =   1968
         Width           =   1164
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "TEL"
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
         Left            =   2310
         TabIndex        =   27
         Top             =   1725
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "------------------------------------- ***DELINQUENT NOTICE*** ------------------------------------- "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   2208
         TabIndex        =   26
         Top             =   96
         Width           =   2412
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Date"
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
         Left            =   2208
         TabIndex        =   25
         Tag             =   "Date is supplied at run time."
         Top             =   624
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
         Left            =   1008
         TabIndex        =   24
         Top             =   2112
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
         Left            =   1008
         TabIndex        =   23
         Top             =   2304
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
         Left            =   1008
         TabIndex        =   22
         Top             =   2496
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
         Left            =   1008
         TabIndex        =   21
         Top             =   2688
         Width           =   1740
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLDlqnTemplate.frx":0F4D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   1008
         TabIndex        =   20
         Top             =   3072
         Width           =   4800
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "We show your account delinquent for the following license code(s):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1008
         TabIndex        =   19
         Top             =   4128
         Width           =   3564
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "Category Code 1"
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
         Left            =   1008
         TabIndex        =   18
         Top             =   4608
         Width           =   1404
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "Category Code 2"
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
         Left            =   1008
         TabIndex        =   17
         Top             =   4800
         Width           =   1404
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "Category Code 3"
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
         Left            =   1008
         TabIndex        =   16
         Top             =   4992
         Width           =   1404
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   "Category Code 4"
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
         Left            =   1008
         TabIndex        =   15
         Top             =   5184
         Width           =   1404
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   "Category Code 5"
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
         Left            =   1008
         TabIndex        =   14
         Top             =   5376
         Width           =   1404
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLDlqnTemplate.frx":1093
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   588
         Left            =   1008
         TabIndex        =   13
         Top             =   5664
         Width           =   4764
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "If payment has been made prior to receiving this notice, please disregard this notice."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1008
         TabIndex        =   12
         Top             =   6384
         Width           =   4668
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "Sincerely,"
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
         Left            =   3792
         TabIndex        =   11
         Top             =   7104
         Width           =   1404
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   675
      Left            =   9375
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "Press this 'Next Notice' button to close this delinquent notice screen and open up the screen for delinquent notice #2."
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
      ButtonDesigner  =   "frmBLDlqnTemplate.frx":115C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   690
      Left            =   9375
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "Press this 'Last Notice' button to close this delinquent notice screen and open up the screen for delinquent notice #3."
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
      ButtonDesigner  =   "frmBLDlqnTemplate.frx":133E
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   9936
      TabIndex        =   33
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
      TabIndex        =   34
      Top             =   4032
      Width           =   2052
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   972
      Left            =   9312
      Top             =   1728
      Width           =   1980
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delinquent Notice #1"
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
      Height          =   732
      Left            =   9696
      TabIndex        =   29
      Top             =   1872
      Width           =   1308
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
Attribute VB_Name = "frmBLDlqnTemplate1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmBLDlqnTemplate1
  DoEvents
  frmBLTownSetup.Show
  frmBLTownSetup.fpcmbDLQNotice.SetFocus
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    frmBLMessageBox.Label1.Top = 800
    frmBLMessageBox.Label1.Height = 1100
    frmBLMessageBox.Label1.Caption = "Penalty amounts and X values will be entered in fields on the delinquent notice printing screen. Category Codes 1 - 5 will also be updated at run time."
    frmBLMessageBox.Label2.Top = 2100
    frmBLMessageBox.Label2.Height = 1500
    frmBLMessageBox.Label2.Caption = "Some of the discretionary values appearing on this page are supplied from the Town Setup screen. If other delinquent notice templates have been used then some of the values here may have carried over from them. PLEASE REVIEW ALL values to make sure they reflect the CURRENT situation."
    frmBLMessageBox.Show vbModal
    fptxtTownOf.ToolTipText = ""
    fptxtAdd.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    fptxtPhone.ToolTipText = ""
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
'    fptxtAdd.ToolTipText = "Enter your town's mailing address here."
'    fptxtCity.ToolTipText = "Enter your town's mailing name here."
'    fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
'    fptxtZip.ToolTipText = "Enter your town's zip code here."
'    fptxtPhone.ToolTipText = "Enter the town's official phone number here."
'    fptxtSigner.ToolTipText = "Enter the signer's name here."
'    fptxtTitle.ToolTipText = "Enter the signer's job title here."
'    cmdHelp.ToolTipText = "If Help is turned on then click to deactivate the informational balloons. If turned off then press to activate instructional balloons."
'    cmdNext.ToolTipText = "Press to move to delinquent notice #2."
'    cmdLast.ToolTipText = "Press to move to delinqueunt notice #3."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = "Press to save the data on this screen."
  End If
    
    
End Sub

Private Sub cmdLast_Click()
  frmBLFreeFormatDlnq.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdNext_Click()
  frmBLDlqnTemplate2a.Show
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
  
  If Len(QPTrim$(fptxtAdd.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town's mailing address for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtAdd.BackColor = &H80FFFF
    fptxtAdd.SetFocus
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
  
  If Len(QPTrim$(fptxtState.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town's state for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtState.BackColor = &H80FFFF
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If Len(QPTrim$(fptxtZip.Text)) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please save the town's zip code for your Delinquent Notice."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtZip.BackColor = &H80FFFF
    fptxtZip.SetFocus
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
    frmBLMessageBoxJr.Label1.Caption = "Please save the title of the signer of your Delinquent Notice"
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtTitle.BackColor = &H80FFFF
    fptxtTitle.SetFocus
    Exit Sub
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec 'other data already saved
    TownRec.DlqAdd1 = QPTrim$(fptxtAdd.Text)
    TownRec.DlqAdminName = QPTrim$(fptxtSigner.Text)
    TownRec.DlqAdminTitle = QPTrim$(fptxtTitle.Text)
    TownRec.DlqCity = QPTrim$(fptxtCity.Text)
    TownRec.DlqPhone = QPTrim$(fptxtPhone.Text)
    TownRec.DlqState = QPTrim$(fptxtState.Text)
    TownRec.DlqTownName = QPTrim$(fptxtTownOf.Text)
    TownRec.DlqZip = QPTrim$(fptxtZip.Text)
    TownRec.DLQNotice = 1
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
    TownRec.DLQNotice = 1
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
    TownRec.DlqAdd1 = QPTrim$(fptxtAdd.Text)
    TownRec.DlqAdminName = QPTrim$(fptxtSigner.Text)
    TownRec.DlqAdminTitle = QPTrim$(fptxtTitle.Text)
    TownRec.DlqCity = QPTrim$(fptxtCity.Text)
    TownRec.DlqPhone = QPTrim$(fptxtPhone.Text)
    TownRec.DlqPhone2 = "" 'not used on this form
    TownRec.DlqFax = "" 'not used on this form
    TownRec.DlqState = QPTrim$(fptxtState.Text)
    TownRec.DlqTownName = QPTrim$(fptxtTownOf.Text)
    TownRec.DlqZip = QPTrim$(fptxtZip.Text)
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
    TownRec.LaserLtr = "N"
    TownRec.GL2Cats = "N"
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
  
  Close THandle
  frmBLSucSave.Label1.Caption = "Your delinquent notice template #1 has been saved."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  
  frmBLTownSetup.fpcmbDLQNotice.Text = "1. PENALTY STANDARD"
  frmBLTownSetup.fpcmdDLQ.Text = "F5 Show Dl&q Notice 1"
  Call cmdExit_Click
  
  MainLog ("Delinquent template #1 saved.")
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDlqnTemplate1", "cmdSave_Click", Erl)
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
  fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
  fptxtAdd.ToolTipText = "Enter your town's mailing address here."
  fptxtCity.ToolTipText = "Enter your town's mailing name here."
  fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
  fptxtZip.ToolTipText = "Enter your town's zip code here."
  fptxtPhone.ToolTipText = "Enter the town's official phone number here."
  fptxtSigner.ToolTipText = "Enter the signer's name here."
  fptxtTitle.ToolTipText = "Enter the signer's job title here."
  cmdHelp.ToolTipText = "If Help is turned on then click to deactivate the informational balloons. If turned off then press to activate instructional balloons."
  cmdNext.ToolTipText = "Press to move to delinquent notice #2."
  cmdLast.ToolTipText = "Press to move to delinqueunt notice #3."
  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
  cmdSave.ToolTipText = "Press to save the data on this screen."
  
  Label7.Caption = " According to our records, your 20XX Business License has not been purchased as of today.  All licenses are now subject to a penalty and will NOT be issued unless the penalty amount is included with your payment.  We realize that you are very busy,  but we would like for you to take the time to purchase this license."
  Label14.Caption = "Please remit your payment  (including penalty) to this office NO later than: Day Of Week, Month XX, 20XX. If you have questions regarding your license, please feel free to contact our office."
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    If Len(QPTrim$(TownRec.DlqAdd1)) = 0 Then
      If Len(QPTrim$(TownRec.TownAdd1)) > 0 Then
        fptxtAdd.Text = QPTrim$(TownRec.TownAdd1)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtAdd1.Text)) > 0 Then
        fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
      End If
    Else
      fptxtAdd.Text = QPTrim$(TownRec.DlqAdd1)
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
    
    If Len(QPTrim$(TownRec.DlqState)) = 0 Then
      If Len(QPTrim$(TownRec.State)) > 0 Then
        fptxtState.Text = QPTrim$(TownRec.State)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtState.Text)) > 0 Then
        fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
      End If
    Else
      fptxtState.Text = QPTrim$(TownRec.DlqState)
    End If
    
    If Len(QPTrim$(TownRec.DlqTownName)) = 0 Then
      If Len(QPTrim$(TownRec.TownName)) > 0 Then
        fptxtTownOf.Text = QPTrim$(TownRec.TownName)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtTownName.Text)) > 0 Then
        fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
      End If
    Else
      fptxtTownOf.Text = QPTrim$(TownRec.DlqTownName)
    End If
    
    If Len(QPTrim$(TownRec.DlqZip)) = 0 Then
      If Len(QPTrim$(TownRec.ZipCode)) > 0 Then
        fptxtZip.Text = QPTrim$(TownRec.ZipCode)
      ElseIf Len(QPTrim$(frmBLTownSetup.fptxtZip.Text)) > 0 Then
        fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
      End If
    Else
      fptxtZip.Text = QPTrim$(TownRec.DlqZip)
    End If
    fptxtSigner.Text = QPTrim$(TownRec.DlqAdminName)
    fptxtTitle.Text = QPTrim$(TownRec.DlqAdminTitle)
  Else
    If QPTrim$(frmBLTownSetup.fptxtTownName.Text) = "" Then
      fptxtTownOf.Text = "Town of YourTown"
    Else
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) = "" Then
      fptxtAdd.Text = "Main St."
    Else
      fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtCity.Text) = "" Then
      fptxtCity.Text = "Your City"
    Else
      fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtState.Text) = "" Then
      fptxtState.Text = "NC"
    Else
      fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtZip.Text) = "" Then
      fptxtZip.Text = "11111-1111"
    Else
      fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtPhone.Text) = "(" Then
      fptxtPhone.Text = "(555)-555-5555"
    Else
      fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
    End If
  
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDlqnTemplate1", "LoadMe", Erl)
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

Private Sub fptxtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAdd.BackColor = -2147483643
End Sub

Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCity.BackColor = -2147483643
End Sub

Private Sub fptxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtPhone.BackColor = -2147483643
End Sub

Private Sub fptxtSigner_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtSigner.BackColor = -2147483643
End Sub

Private Sub fptxtState_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtState.BackColor = -2147483643
End Sub

Private Sub fptxtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTitle.BackColor = -2147483643
End Sub

Private Sub fptxtTownOf_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownOf.BackColor = -2147483643
End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtZip.BackColor = -2147483643
End Sub
Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  Me.PrintForm
  MainLog ("Delinquent Notice # 1:  screen printed.")
End Sub

