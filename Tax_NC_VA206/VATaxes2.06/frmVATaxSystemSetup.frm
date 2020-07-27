VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxSystemSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax System Setup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "frmVATaxSystemSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6855
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   11535
      _Version        =   196609
      _ExtentX        =   20346
      _ExtentY        =   12091
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   2
      TabCount        =   2
      Tab             =   1
      ShowFocusRect   =   0   'False
      ActiveTabBold   =   0   'False
      GrayAreaColor   =   13684944
      OffsetFromClientTop=   -1  'True
      HighestPrecedence=   1
      ShowEarMark     =   0   'False
      DataFormat      =   ""
      AutoSizeChildren=   3
      BookCornerGuardWidth=   60
      BookCornerGuardLength=   360
      DrawFocusRect   =   1
      DataField       =   ""
      TabCaption      =   "frmVATaxSystemSetup.frx":08CA
      PageEarMarkPictureNext=   "frmVATaxSystemSetup.frx":0B32
      PageEarMarkPicturePrev=   "frmVATaxSystemSetup.frx":0B4E
      EarMarkPictureNext=   "frmVATaxSystemSetup.frx":0B6A
      EarMarkPicturePrev=   "frmVATaxSystemSetup.frx":0B86
      Begin ImpproLib.vaImprint vaImprint6 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   22
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":0BA2
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   1
            Left            =   2880
            TabIndex        =   23
            Top             =   1536
            Width           =   5400
            _Version        =   196608
            _ExtentX        =   9525
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
         Begin EditLib.fpCurrency fptxtE 
            Height          =   348
            Index           =   1
            Left            =   8448
            TabIndex        =   24
            ToolTipText     =   "Enter the Amount for this earnings code."
            Top             =   1536
            Width           =   1212
            _Version        =   196608
            _ExtentX        =   2138
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtD 
            Height          =   348
            Index           =   1
            Left            =   624
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1536
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
            _ExtentY        =   614
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
            BackColor       =   -2147483634
            ForeColor       =   -2147483630
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
            BorderColor     =   -2147483631
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
            InvalidColor    =   -2147483630
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
            ButtonColor     =   -2147483634
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fptxtD 
            Height          =   348
            Index           =   5
            Left            =   624
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   3456
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtD 
            Height          =   348
            Index           =   4
            Left            =   624
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2976
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtD 
            Height          =   348
            Index           =   3
            Left            =   624
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2496
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtD 
            Height          =   348
            Index           =   2
            Left            =   624
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   2016
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
            _ExtentY        =   614
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
         Begin EditLib.fpCurrency fptxtE 
            Height          =   348
            Index           =   2
            Left            =   8448
            TabIndex        =   30
            ToolTipText     =   "Enter the Amount for this earnings code."
            Top             =   2016
            Width           =   1212
            _Version        =   196608
            _ExtentX        =   2138
            _ExtentY        =   614
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
         Begin EditLib.fpCurrency fptxtE 
            Height          =   348
            Index           =   3
            Left            =   8448
            TabIndex        =   31
            ToolTipText     =   "Enter the Amount for this earnings code."
            Top             =   2496
            Width           =   1212
            _Version        =   196608
            _ExtentX        =   2138
            _ExtentY        =   614
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
         Begin EditLib.fpCurrency fptxtE 
            Height          =   348
            Index           =   4
            Left            =   8448
            TabIndex        =   32
            ToolTipText     =   "Enter the Amount for this earnings code."
            Top             =   2976
            Width           =   1212
            _Version        =   196608
            _ExtentX        =   2138
            _ExtentY        =   614
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
         Begin EditLib.fpCurrency fptxtE 
            Height          =   348
            Index           =   5
            Left            =   8448
            TabIndex        =   33
            ToolTipText     =   "Enter the Amount for this earnings code."
            Top             =   3456
            Width           =   1212
            _Version        =   196608
            _ExtentX        =   2138
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   2
            Left            =   2880
            TabIndex        =   34
            Top             =   2016
            Width           =   5400
            _Version        =   196608
            _ExtentX        =   9525
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
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   3
            Left            =   2880
            TabIndex        =   35
            Top             =   2496
            Width           =   5400
            _Version        =   196608
            _ExtentX        =   9525
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
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   4
            Left            =   2880
            TabIndex        =   36
            Top             =   2976
            Width           =   5400
            _Version        =   196608
            _ExtentX        =   9525
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
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   5
            Left            =   2880
            TabIndex        =   37
            Top             =   3456
            Width           =   5400
            _Version        =   196608
            _ExtentX        =   9525
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
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   1056
            TabIndex        =   40
            Top             =   1104
            Width           =   1308
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number"
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
            Left            =   4848
            TabIndex        =   39
            Top             =   1104
            Width           =   1884
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   8592
            TabIndex        =   38
            Top             =   1104
            Width           =   972
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
      End
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   41
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":0BBE
         Begin LpLib.fpCombo fpcomboEIC 
            Height          =   405
            Left            =   4710
            TabIndex        =   42
            ToolTipText     =   "Select this employee's EIC code from pick list."
            Top             =   3840
            Width           =   2700
            _Version        =   196608
            _ExtentX        =   4762
            _ExtentY        =   714
            Text            =   ""
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   2
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
            ColumnEdit      =   0
            ColumnBound     =   0
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   2
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":0BDA
         End
         Begin LpLib.fpCombo fpcomboMedX 
            Height          =   405
            Left            =   6000
            TabIndex        =   43
            ToolTipText     =   "Answer Y if this employee is exempt from medicare tax withholding."
            Top             =   3270
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":1079
         End
         Begin LpLib.fpCombo fpcomboSocX 
            Height          =   405
            Left            =   6330
            TabIndex        =   44
            ToolTipText     =   "Answer Y if this employee is exempt from SS tax withholding."
            Top             =   2685
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":14C0
         End
         Begin LpLib.fpCombo fpcomboFedStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   45
            ToolTipText     =   "Select the employee's Federal filing status."
            Top             =   1530
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
            _ExtentY        =   714
            Text            =   ""
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   3
            ColumnSearch    =   -1
            ColumnWidthScale=   2
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
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
            ColumnEdit      =   0
            ColumnBound     =   0
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":1907
         End
         Begin LpLib.fpCombo fpcomboStateStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   46
            ToolTipText     =   "Select the employee's State filing status."
            Top             =   2115
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
            _ExtentY        =   714
            Text            =   ""
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   3
            ColumnSearch    =   -1
            ColumnWidthScale=   2
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
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
            ColumnEdit      =   0
            ColumnBound     =   0
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":1DA6
         End
         Begin LpLib.fpCombo fpcomboStateAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   47
            ToolTipText     =   "Override tax calc and enter an ""A"" for an amount or ""P"" for percentage."
            Top             =   2115
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":2245
         End
         Begin LpLib.fpCombo fpcomboFedAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   48
            ToolTipText     =   "Override tax calc and enter an ""A"" for an amount or ""P"" for percentage."
            Top             =   1530
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":268C
         End
         Begin LpLib.fpCombo fpcomboStateX 
            Height          =   405
            Left            =   1935
            TabIndex        =   49
            ToolTipText     =   "Answer Y if this employee is exempt from fed tax withholding."
            Top             =   2115
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":2AD3
         End
         Begin LpLib.fpCombo fpcomboFedX 
            Height          =   405
            Left            =   1935
            TabIndex        =   50
            ToolTipText     =   "Answer Y if this employee is exempt from fed tax withholding."
            Top             =   1530
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":2F1A
         End
         Begin EditLib.fpCurrency fptxtAddWHFed 
            Height          =   396
            Left            =   8352
            TabIndex        =   51
            Top             =   1536
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
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
         Begin EditLib.fpText fptxtAllowNumState 
            Height          =   396
            Left            =   7104
            TabIndex        =   52
            ToolTipText     =   "Enter the Number of Allowances Claimed by this Employee here."
            Top             =   2112
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
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
            MaxLength       =   2
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
         Begin EditLib.fpText fptxtAllowNumFed 
            Height          =   396
            Left            =   7104
            TabIndex        =   53
            ToolTipText     =   "Enter the Number of Allowances Claimed by this Employee here."
            Top             =   1536
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
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
            MaxLength       =   2
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
         Begin EditLib.fpCurrency fptxtAddWHState 
            Height          =   396
            Left            =   8352
            TabIndex        =   54
            Top             =   2112
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
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
         Begin EditLib.fpText fptxtFedFig 
            Height          =   384
            Left            =   4464
            TabIndex        =   55
            ToolTipText     =   "Enter the Number of Allowances Claimed by this Employee here."
            Top             =   1536
            Width           =   1068
            _Version        =   196608
            _ExtentX        =   1884
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
            CharValidationText=   "0 1 2 3 4 5 6 7 8 9 ."
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
         Begin EditLib.fpText fptxtStateFig 
            Height          =   384
            Left            =   4464
            TabIndex        =   56
            ToolTipText     =   "Enter the Number of Allowances Claimed by this Employee here."
            Top             =   2112
            Width           =   1068
            _Version        =   196608
            _ExtentX        =   1884
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
            CharValidationText=   "0 1 2 3 4 5 6 7 8 9 ."
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
         Begin VB.Shape Shape5 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Federal"
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
            Height          =   396
            Left            =   672
            TabIndex        =   69
            Top             =   1680
            Width           =   828
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   912
            TabIndex        =   68
            Top             =   2208
            Width           =   684
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Exempt*"
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
            Left            =   1872
            TabIndex        =   67
            Top             =   1056
            Width           =   1068
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Fixed"
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
            Left            =   4032
            TabIndex        =   66
            Top             =   768
            Width           =   636
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Amt/Pct"
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
            Left            =   3264
            TabIndex        =   65
            Top             =   1056
            Width           =   972
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Figure"
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
            Height          =   396
            Left            =   4656
            TabIndex        =   64
            Top             =   1056
            Width           =   732
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Status*"
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
            Left            =   5760
            TabIndex        =   63
            Top             =   1056
            Width           =   828
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Allowances*"
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
            Left            =   6816
            TabIndex        =   62
            Top             =   1056
            Width           =   1356
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Additional  "
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
            Height          =   288
            Left            =   8256
            TabIndex        =   61
            Top             =   768
            Width           =   1188
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Social Security Exempt?*"
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
            Left            =   3504
            TabIndex        =   60
            Top             =   2784
            Width           =   2700
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicare Exempt?*"
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
            Left            =   3744
            TabIndex        =   59
            Top             =   3360
            Width           =   2124
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "EIC Code*"
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
            Left            =   3312
            TabIndex        =   58
            Top             =   3936
            Width           =   1212
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "W/H Amt"
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
            Left            =   8304
            TabIndex        =   57
            Top             =   1056
            Width           =   1020
         End
      End
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   70
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":3361
         Begin LpLib.fpCombo fpcomboPayType 
            Height          =   405
            Left            =   2685
            TabIndex        =   71
            ToolTipText     =   "Select the Employee's Pay Type from the pick list."
            Top             =   2310
            Width           =   1935
            _Version        =   196608
            _ExtentX        =   3413
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":337D
         End
         Begin LpLib.fpCombo fpcomboStatus 
            Height          =   405
            Left            =   2685
            TabIndex        =   72
            ToolTipText     =   "Select the Employee's Employment Status from the pick list."
            Top             =   1725
            Width           =   1935
            _Version        =   196608
            _ExtentX        =   3413
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":37C4
         End
         Begin LpLib.fpCombo fpcomboFreq 
            Height          =   405
            Left            =   2685
            TabIndex        =   73
            ToolTipText     =   "Select the Employee's Pay Frequency from the pick list."
            Top             =   2880
            Width           =   1935
            _Version        =   196608
            _ExtentX        =   3413
            _ExtentY        =   714
            Text            =   ""
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
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   200
            EditMarginLeft  =   1
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   1
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":3C0B
         End
         Begin EditLib.fpCurrency fptxtRate 
            Height          =   450
            Left            =   7830
            TabIndex        =   74
            Top             =   1155
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1968
            _ExtentY        =   783
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
         Begin EditLib.fpText fptxtBenefitPct 
            Height          =   450
            Left            =   2685
            TabIndex        =   75
            ToolTipText     =   "Enter the Employee's Leave Benefit Percentage."
            Top             =   3450
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1968
            _ExtentY        =   783
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
         Begin EditLib.fpText txtTitle 
            Height          =   450
            Left            =   3075
            TabIndex        =   76
            ToolTipText     =   "Enter the Employee's Job Title here."
            Top             =   450
            Width           =   5250
            _Version        =   196608
            _ExtentX        =   9250
            _ExtentY        =   783
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
            MaxLength       =   28
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
         Begin EditLib.fpText fptxtWCCode 
            Height          =   450
            Left            =   2730
            TabIndex        =   77
            ToolTipText     =   "Enter the Employee's Worker's Compensation Classification here."
            Top             =   1110
            Width           =   1020
            _Version        =   196608
            _ExtentX        =   1799
            _ExtentY        =   783
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
            MaxLength       =   12
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
         Begin EditLib.fpDateTime fpMaskNext 
            Height          =   375
            Left            =   7830
            TabIndex        =   78
            ToolTipText     =   "If available, enter the date for this employee's next review."
            Top             =   2880
            Width           =   1695
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
            Text            =   "10-01-2001"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm-dd-yyyy"
            DateMax         =   "20350101"
            DateMin         =   "19200101"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "19800101"
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
         Begin EditLib.fpDateTime fpMaskHire 
            Height          =   375
            Left            =   7830
            TabIndex        =   79
            ToolTipText     =   "Enter the date this employee was hired."
            Top             =   2310
            Width           =   1695
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
            Text            =   "10-01-2001"
            DateCalcMethod  =   1
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm-dd-yyyy"
            DateMax         =   "20350101"
            DateMin         =   "19200101"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "19200101"
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
         Begin EditLib.fpDateTime fpMaskTerm 
            Height          =   375
            Left            =   7830
            TabIndex        =   80
            ToolTipText     =   "If terminated, enter the date this employee was terminated."
            Top             =   3450
            Width           =   1695
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
            AutoAdvance     =   0   'False
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
            Text            =   "10-01-2001"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm-dd-yyyy"
            DateMax         =   "20350101"
            DateMin         =   "19200101"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "19800101"
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
         Begin EditLib.fpCurrency fptxtOTRate 
            Height          =   450
            Left            =   7830
            TabIndex        =   81
            Top             =   1725
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1968
            _ExtentY        =   783
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
         Begin EditLib.fpText fptxtComment 
            Height          =   450
            Left            =   3075
            TabIndex        =   82
            ToolTipText     =   "Up to 25 characters are available for a comment."
            Top             =   4200
            Width           =   5250
            _Version        =   196608
            _ExtentX        =   9260
            _ExtentY        =   794
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
            MaxLength       =   25
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
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
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
            Left            =   2205
            TabIndex        =   94
            Top             =   645
            Width           =   690
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "W/C Code*"
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
            Left            =   1155
            TabIndex        =   93
            Top             =   1245
            Width           =   1260
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Status*"
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
            Left            =   1200
            TabIndex        =   92
            Top             =   1830
            Width           =   1065
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Type*"
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
            Left            =   1200
            TabIndex        =   91
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency*"
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
            Height          =   390
            Left            =   1200
            TabIndex        =   90
            Top             =   2970
            Width           =   1455
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Pct*"
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
            Left            =   1200
            TabIndex        =   89
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate*"
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
            Left            =   5670
            TabIndex        =   88
            Top             =   1290
            Width           =   870
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "OT Rate"
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
            Left            =   5670
            TabIndex        =   87
            Top             =   1830
            Width           =   1215
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Hire Date"
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
            Left            =   5715
            TabIndex        =   86
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Next Review "
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
            Left            =   5715
            TabIndex        =   85
            Top             =   3030
            Width           =   1500
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Termination  Date"
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
            Left            =   5715
            TabIndex        =   84
            Top             =   3600
            Width           =   2070
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
         Begin VB.Label Label77 
            BackStyle       =   0  'Transparent
            Caption         =   "Comment"
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
            Left            =   1680
            TabIndex        =   83
            Top             =   4320
            Width           =   1185
         End
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   6330
         Left            =   -26310
         TabIndex        =   95
         Top             =   -21690
         Width           =   11265
         _Version        =   196609
         _ExtentX        =   19870
         _ExtentY        =   11165
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":4052
         Begin LpLib.fpList fpListTownships 
            Height          =   540
            Left            =   8760
            TabIndex        =   8
            Top             =   1560
            Width           =   2295
            _Version        =   196608
            _ExtentX        =   4048
            _ExtentY        =   952
            TextAlias       =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            Columns         =   0
            Sorted          =   0
            LineWidth       =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   -1
            ColumnWidthScale=   2
            RowHeight       =   -1
            MultiSelect     =   0
            WrapList        =   0   'False
            WrapWidth       =   0
            SelMax          =   -1
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
            DataField       =   ""
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":406E
         End
         Begin LpLib.fpCombo fpcmbMinOptions 
            Height          =   330
            Left            =   6360
            TabIndex        =   18
            Top             =   5760
            Width           =   4500
            _Version        =   196608
            _ExtentX        =   7937
            _ExtentY        =   582
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
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
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":444A
         End
         Begin LpLib.fpCombo fpcmbStateOfTax 
            Height          =   390
            Left            =   7635
            TabIndex        =   6
            Tag             =   $"frmVATaxSystemSetup.frx":4891
            Top             =   1920
            Width           =   900
            _Version        =   196608
            _ExtentX        =   1587
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":49F9
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdAddTownship 
            Height          =   375
            Left            =   8760
            TabIndex        =   96
            Top             =   960
            Width           =   2295
            _Version        =   131072
            _ExtentX        =   4048
            _ExtentY        =   661
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":4E40
         End
         Begin EditLib.fpCurrency fptxtMinTaxAmt 
            Height          =   375
            Left            =   2640
            TabIndex        =   17
            Top             =   5760
            Width           =   1095
            _Version        =   196608
            _ExtentX        =   1931
            _ExtentY        =   661
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
         Begin EditLib.fpDoubleSingle fptxtCurrYrRIntRate 
            Height          =   372
            Left            =   3720
            TabIndex        =   9
            ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
            Top             =   2652
            Width           =   1092
            _Version        =   196608
            _ExtentX        =   1926
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin EditLib.fpText fptxtAdd1 
            Height          =   390
            Left            =   2880
            TabIndex        =   1
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   720
            Width           =   5655
            _Version        =   196608
            _ExtentX        =   9975
            _ExtentY        =   688
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
            CharValidationText=   ""
            MaxLength       =   35
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
         Begin EditLib.fpText fptxtAdd2 
            Height          =   390
            Left            =   2880
            TabIndex        =   2
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   1110
            Width           =   5655
            _Version        =   196608
            _ExtentX        =   9975
            _ExtentY        =   688
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
            CharValidationText=   ""
            MaxLength       =   35
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
            Height          =   390
            Left            =   2880
            TabIndex        =   3
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   1500
            Width           =   5655
            _Version        =   196608
            _ExtentX        =   9975
            _ExtentY        =   688
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
         Begin EditLib.fpText fptxtState 
            Height          =   390
            Left            =   2880
            TabIndex        =   4
            Tag             =   "Enter the state the town is in. Use the generally accepted two character upper case abbreviation (North Carolina = NC)."
            Top             =   1890
            Width           =   585
            _Version        =   196608
            _ExtentX        =   1032
            _ExtentY        =   688
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
         Begin EditLib.fpMask fptxtZip 
            Height          =   390
            Left            =   4710
            TabIndex        =   5
            Tag             =   "Enter either a five digit or nine digit zip code for the town in this field."
            Top             =   1890
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
            _ExtentY        =   688
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
         Begin EditLib.fpDoubleSingle fptxtPastYearIntRate 
            Height          =   372
            Left            =   9648
            TabIndex        =   13
            ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
            Top             =   3240
            Width           =   1092
            _Version        =   196608
            _ExtentX        =   1926
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin EditLib.fpText fptxtCustOptSrch 
            Height          =   390
            Left            =   1800
            TabIndex        =   14
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   4050
            Width           =   2055
            _Version        =   196608
            _ExtentX        =   3625
            _ExtentY        =   688
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
         Begin EditLib.fpText fptxtPropOptSrch 
            Height          =   390
            Left            =   5280
            TabIndex        =   15
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   4050
            Width           =   2055
            _Version        =   196608
            _ExtentX        =   3625
            _ExtentY        =   688
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
         Begin EditLib.fpDateTime fptxtCurrRYear 
            Height          =   348
            Left            =   7560
            TabIndex        =   11
            Top             =   2750
            Width           =   1020
            _Version        =   196608
            _ExtentX        =   1799
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
            ControlType     =   0
            Text            =   "2018"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "yyyy"
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
         Begin EditLib.fpText fptxtTownShipName 
            Height          =   390
            Left            =   8760
            TabIndex        =   7
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   480
            Width           =   2295
            _Version        =   196608
            _ExtentX        =   4048
            _ExtentY        =   688
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
         Begin EditLib.fpText fptxtNameOfTaxAuth 
            Height          =   390
            Left            =   2880
            TabIndex        =   0
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   330
            Width           =   5655
            _Version        =   196608
            _ExtentX        =   9975
            _ExtentY        =   688
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
            CharValidationText=   ""
            MaxLength       =   35
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
         Begin EditLib.fpDoubleSingle fptxtCurrYrPIntRate 
            Height          =   372
            Left            =   3720
            TabIndex        =   10
            ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
            Top             =   3240
            Width           =   1092
            _Version        =   196608
            _ExtentX        =   1926
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin fpBtnAtlLibCtl.fpBtn cmdRealPctSetup 
            Height          =   324
            Left            =   4920
            TabIndex        =   236
            TabStop         =   0   'False
            Top             =   2688
            Width           =   1800
            _Version        =   131072
            _ExtentX        =   3175
            _ExtentY        =   572
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":5024
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdPersTaxSetup 
            Height          =   324
            Left            =   4920
            TabIndex        =   237
            TabStop         =   0   'False
            Top             =   3276
            Width           =   1800
            _Version        =   131072
            _ExtentX        =   3175
            _ExtentY        =   572
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":5205
         End
         Begin EditLib.fpDateTime fptxtCurrPYear 
            Height          =   348
            Left            =   9960
            TabIndex        =   12
            Top             =   2750
            Width           =   1020
            _Version        =   196608
            _ExtentX        =   1799
            _ExtentY        =   614
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
            Text            =   "2018"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "yyyy"
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
         Begin EditLib.fpText fptxtPersOptSrch 
            Height          =   390
            Left            =   8880
            TabIndex        =   16
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   4050
            Width           =   2055
            _Version        =   196608
            _ExtentX        =   3625
            _ExtentY        =   688
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
         Begin VB.Label Label105 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "For Personal:"
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
            Height          =   270
            Left            =   7440
            TabIndex        =   241
            Top             =   4170
            Width           =   1380
         End
         Begin VB.Line Line14 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   6880
            X2              =   11160
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   6885
            X2              =   6885
            Y1              =   2400
            Y2              =   3180
         End
         Begin VB.Label Label103 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Personal:"
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
            Left            =   8760
            TabIndex        =   239
            Top             =   2846
            Width           =   1020
         End
         Begin VB.Label Label102 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Current Tax Year:"
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
            Left            =   8160
            TabIndex        =   238
            Top             =   2440
            Width           =   1740
         End
         Begin VB.Label Label96 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year Personal Interest Rate:"
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
            Left            =   120
            TabIndex        =   235
            Top             =   3348
            Width           =   3540
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   1745
            Index           =   0
            Left            =   120
            Top             =   4530
            Width           =   11055
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Minimum Tax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   117
            Top             =   4545
            Width           =   2340
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Tax Options:"
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
            Height          =   270
            Left            =   4080
            TabIndex        =   116
            Top             =   5850
            Width           =   2220
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Tax Amount:"
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
            Height          =   270
            Left            =   360
            TabIndex        =   115
            Top             =   5850
            Width           =   2220
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmVATaxSystemSetup.frx":53EA
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
            Height          =   855
            Left            =   240
            TabIndex        =   114
            Top             =   4875
            Width           =   10620
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Past Year Interest Rate:"
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
            Left            =   7116
            TabIndex        =   113
            Top             =   3360
            Width           =   2340
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year Real Interest Rate:"
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
            Left            =   120
            TabIndex        =   112
            Top             =   2760
            Width           =   3540
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*State of Tax:"
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
            Height          =   270
            Left            =   6000
            TabIndex        =   111
            Top             =   1995
            Width           =   1500
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Zip Code:"
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
            Height          =   270
            Left            =   3600
            TabIndex        =   110
            Top             =   1995
            Width           =   1020
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*State:"
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
            Height          =   270
            Left            =   1680
            TabIndex        =   109
            Top             =   1995
            Width           =   1020
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*City:"
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
            Height          =   270
            Left            =   1800
            TabIndex        =   108
            Top             =   1605
            Width           =   900
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2:"
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
            Height          =   270
            Left            =   240
            TabIndex        =   107
            Top             =   1230
            Width           =   2460
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Address 1:"
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
            Height          =   270
            Left            =   240
            TabIndex        =   106
            Top             =   825
            Width           =   2460
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Taxing Authority Name:"
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
            Height          =   270
            Left            =   240
            TabIndex        =   105
            Top             =   435
            Width           =   2460
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   2145
            Left            =   120
            Top             =   2400
            Width           =   11055
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   2292
            Left            =   120
            Top             =   120
            Width           =   11052
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Interest Rate Applications"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   120
            TabIndex        =   104
            Top             =   2400
            Width           =   3060
         End
         Begin VB.Label Label78 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Billing Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   103
            Top             =   120
            Width           =   2340
         End
         Begin VB.Label Label76 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Optional Search Fields"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   120
            TabIndex        =   102
            Top             =   3720
            Width           =   2700
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "For Customer:"
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
            Height          =   270
            Left            =   360
            TabIndex        =   101
            Top             =   4170
            Width           =   1380
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "For Abstract:"
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
            Height          =   270
            Left            =   3960
            TabIndex        =   100
            Top             =   4170
            Width           =   1260
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   11160
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Real:"
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
            Left            =   6720
            TabIndex        =   99
            Top             =   2846
            Width           =   660
         End
         Begin VB.Label Label86 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Township"
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
            Height          =   270
            Left            =   8880
            TabIndex        =   98
            Top             =   240
            Width           =   2100
         End
         Begin VB.Line Line7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   8640
            X2              =   8640
            Y1              =   120
            Y2              =   2400
         End
         Begin VB.Label Label91 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Township List"
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
            Height          =   270
            Left            =   8880
            TabIndex        =   97
            Top             =   1320
            Width           =   2100
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   6330
         Left            =   150
         TabIndex        =   118
         Top             =   390
         Width           =   11265
         _Version        =   196609
         _ExtentX        =   19870
         _ExtentY        =   11165
         _StockProps     =   70
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":5540
         Begin LpLib.fpCombo fpcmbPPTRAYN 
            Height          =   390
            Left            =   4980
            TabIndex        =   131
            Top             =   3120
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":555C
         End
         Begin LpLib.fpCombo fpcmbMultiYear 
            Height          =   390
            Left            =   8760
            TabIndex        =   127
            Tag             =   $"frmVATaxSystemSetup.frx":59A3
            Top             =   1800
            Width           =   780
            _Version        =   196608
            _ExtentX        =   1376
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":5B0B
         End
         Begin LpLib.fpCombo fpcmbNoInterYN 
            Height          =   390
            Left            =   10275
            TabIndex        =   125
            Tag             =   $"frmVATaxSystemSetup.frx":5F52
            Top             =   1335
            Width           =   780
            _Version        =   196608
            _ExtentX        =   1376
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":60BA
         End
         Begin LpLib.fpCombo fpcmbLateFormat 
            Height          =   375
            Left            =   2115
            TabIndex        =   124
            Tag             =   $"frmVATaxSystemSetup.frx":6501
            ToolTipText     =   "Enter the desired late tax bill format. Press the 'Show Late Tax Bill' button to see a likeness of the selected tax late bill."
            Top             =   1335
            Width           =   2580
            _Version        =   196608
            _ExtentX        =   4551
            _ExtentY        =   661
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":6669
         End
         Begin LpLib.fpCombo fpcmbAcctMeth 
            Height          =   390
            Left            =   8715
            TabIndex        =   123
            ToolTipText     =   "Enter the accounting method to be used for tax billing."
            Top             =   870
            Width           =   2340
            _Version        =   196608
            _ExtentX        =   4128
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":6AB0
         End
         Begin LpLib.fpCombo fpcmbTaxBillFormat 
            Height          =   375
            Left            =   2115
            TabIndex        =   122
            Top             =   870
            Width           =   2580
            _Version        =   196608
            _ExtentX        =   4551
            _ExtentY        =   661
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":6EF7
         End
         Begin LpLib.fpCombo fpcmbCentDepYN 
            Height          =   390
            Left            =   2640
            TabIndex        =   119
            ToolTipText     =   "If the town uses the Central Depository type of accounting then enter 'Yes' here."
            Top             =   360
            Width           =   900
            _Version        =   196608
            _ExtentX        =   1587
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":733E
         End
         Begin LpLib.fpCombo fpcmbCyclesYN 
            Height          =   390
            Left            =   10320
            TabIndex        =   133
            Top             =   3240
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":7785
         End
         Begin LpLib.fpCombo fpcmbCountyYN 
            Height          =   390
            Left            =   10320
            TabIndex        =   134
            Top             =   3720
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":7BCC
         End
         Begin LpLib.fpCombo fpcmbRPSplitYN 
            Height          =   375
            Left            =   10320
            TabIndex        =   135
            Top             =   4200
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
            _ExtentY        =   661
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
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
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmVATaxSystemSetup.frx":8013
         End
         Begin EditLib.fpCurrency fpCurrMaxVehAmt 
            Height          =   372
            Left            =   960
            TabIndex        =   129
            Top             =   3516
            Width           =   1572
            _Version        =   196608
            _ExtentX        =   2773
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
         Begin EditLib.fpDoubleSingle fptxtDiscRPct 
            Height          =   324
            Left            =   9168
            TabIndex        =   136
            ToolTipText     =   $"frmVATaxSystemSetup.frx":845A
            Top             =   5136
            Width           =   1140
            _Version        =   196608
            _ExtentX        =   2011
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin EditLib.fpText fptxtOverPayGL 
            Height          =   396
            Left            =   3000
            TabIndex        =   126
            Top             =   1800
            Width           =   2172
            _Version        =   196608
            _ExtentX        =   3836
            _ExtentY        =   688
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
            CharValidationText=   ""
            MaxLength       =   35
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
         Begin fpBtnAtlLibCtl.fpBtn cmdTaxBill 
            Height          =   324
            Left            =   4800
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   876
            Width           =   1908
            _Version        =   131072
            _ExtentX        =   3365
            _ExtentY        =   572
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":84E5
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdLateBill 
            Height          =   330
            Left            =   4800
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   1335
            Width           =   1905
            _Version        =   131072
            _ExtentX        =   3360
            _ExtentY        =   582
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":86C9
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdCycle 
            Height          =   345
            Left            =   6120
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   3270
            Width           =   1710
            _Version        =   131072
            _ExtentX        =   3016
            _ExtentY        =   609
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":88AE
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdCounty 
            Height          =   345
            Left            =   6120
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   3750
            Width           =   1710
            _Version        =   131072
            _ExtentX        =   3016
            _ExtentY        =   609
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":8A91
         End
         Begin EditLib.fpText fptxtCentCash 
            Height          =   390
            Left            =   5160
            TabIndex        =   120
            ToolTipText     =   "This field is only enabled if 'Central Depository Y/N?' is 'Yes'."
            Top             =   360
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
            _ExtentY        =   688
            Enabled         =   0   'False
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
            CharValidationText=   ""
            MaxLength       =   35
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
         Begin EditLib.fpText fptxtCentSub 
            Height          =   390
            Left            =   8880
            TabIndex        =   121
            ToolTipText     =   "This field is only enabled if 'Central Depository Y/N?' is 'Yes'."
            Top             =   360
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
            _ExtentY        =   688
            Enabled         =   0   'False
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
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
            CharValidationText=   ""
            MaxLength       =   35
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
         Begin fpBtnAtlLibCtl.fpBtn cmdRealClass 
            Height          =   372
            Left            =   960
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   5556
            Width           =   3828
            _Version        =   131072
            _ExtentX        =   6752
            _ExtentY        =   656
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":8C70
         End
         Begin EditLib.fpDateTime fptxtLawChngDate 
            Height          =   348
            Left            =   6720
            TabIndex        =   128
            Top             =   2400
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   614
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
            Text            =   "10/18/2005"
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
         Begin EditLib.fpCurrency fpCurrMinVehAmt 
            Height          =   372
            Left            =   960
            TabIndex        =   130
            Top             =   3960
            Width           =   1572
            _Version        =   196608
            _ExtentX        =   2773
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
         Begin EditLib.fpDoubleSingle fptxtDiscPPct 
            Height          =   324
            Left            =   9168
            TabIndex        =   137
            ToolTipText     =   $"frmVATaxSystemSetup.frx":8E65
            Top             =   5520
            Width           =   1140
            _Version        =   196608
            _ExtentX        =   2011
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin fpBtnAtlLibCtl.fpBtn cmdRevSetUp 
            Height          =   360
            Left            =   1320
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   5010
            Width           =   3105
            _Version        =   131072
            _ExtentX        =   5477
            _ExtentY        =   635
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
            ButtonDesigner  =   "frmVATaxSystemSetup.frx":8EF0
         End
         Begin EditLib.fpDoubleSingle fpDSPPTRADisc 
            Height          =   372
            Left            =   3960
            TabIndex        =   132
            ToolTipText     =   $"frmVATaxSystemSetup.frx":90D0
            Top             =   3960
            Width           =   1140
            _Version        =   196608
            _ExtentX        =   2011
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "100"
            MinValue        =   "0"
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
         Begin VB.Label Label104 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   5160
            TabIndex        =   240
            Top             =   4080
            Width           =   300
         End
         Begin VB.Line Line13 
            BorderColor     =   &H0080FFFF&
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   11160
            Y1              =   2350
            Y2              =   2350
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0080FFFF&
            BorderStyle     =   3  'Dot
            X1              =   3120
            X2              =   3120
            Y1              =   2880
            Y2              =   4560
         End
         Begin VB.Label Label92 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   4680
            TabIndex        =   234
            Top             =   4056
            Width           =   300
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PPTRA Discount Pct"
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
            Left            =   3600
            TabIndex        =   233
            Top             =   3636
            Width           =   1980
         End
         Begin VB.Label Label95 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PPTRA Y/N?:"
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
            Left            =   3360
            TabIndex        =   232
            Top             =   3240
            Width           =   1380
         End
         Begin VB.Label Label101 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " Vehicle Value Amt"
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
            Left            =   720
            TabIndex        =   231
            Top             =   3204
            Width           =   1980
         End
         Begin VB.Label Label100 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "PPTRA Settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   120
            TabIndex        =   230
            Top             =   2880
            Width           =   2148
         End
         Begin VB.Label Label99 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Bill Type Discounts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   6000
            TabIndex        =   228
            Top             =   4800
            Width           =   2508
         End
         Begin VB.Line Line12 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   11160
            X2              =   6000
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Label Label98 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   10368
            TabIndex        =   227
            Top             =   5592
            Width           =   300
         End
         Begin VB.Label Label97 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Personal Principal Discount:"
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
            Left            =   6480
            TabIndex        =   226
            Top             =   5604
            Width           =   2628
         End
         Begin VB.Line Line11 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   11160
            X2              =   6000
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   11160
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
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
            Left            =   360
            TabIndex        =   225
            Top             =   4080
            Width           =   540
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   6000
            X2              =   120
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label Label94 
            BackStyle       =   0  'Transparent
            Caption         =   "Date the Delinquent/Discount Reg Changes:"
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
            Left            =   2400
            TabIndex        =   224
            Top             =   2496
            Width           =   4260
         End
         Begin VB.Label Label93 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Max:"
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
            Left            =   360
            TabIndex        =   223
            Top             =   3636
            Width           =   540
         End
         Begin VB.Line Line10 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   11160
            X2              =   6000
            Y1              =   6096
            Y2              =   6096
         End
         Begin VB.Line Line9 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   11160
            X2              =   11160
            Y1              =   120
            Y2              =   6100
         End
         Begin VB.Line Line8 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   120
            Y1              =   120
            Y2              =   2880
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Multi-Year Billing:"
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
            Left            =   6960
            TabIndex        =   222
            Top             =   1896
            Width           =   1740
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Central Depository Y/N?"
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
            Left            =   240
            TabIndex        =   158
            Top             =   468
            Width           =   2340
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Tax Bill Format:"
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
            Height          =   270
            Left            =   240
            TabIndex        =   157
            Top             =   945
            Width           =   1740
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Accting Method:"
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
            Height          =   270
            Left            =   6840
            TabIndex        =   156
            Top             =   945
            Width           =   1740
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*Late Bill Format:"
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
            Height          =   270
            Left            =   225
            TabIndex        =   155
            Top             =   1425
            Width           =   1740
         End
         Begin VB.Label Label72 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Accounting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   154
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Overpayment G/L #:"
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
            Left            =   960
            TabIndex        =   153
            Top             =   1920
            Width           =   1980
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Pay Sequence/Opt Revenue Setup/Class Setup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   120
            TabIndex        =   152
            Top             =   4560
            Width           =   4908
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Real Principal Discount:"
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
            Left            =   6912
            TabIndex        =   151
            Top             =   5220
            Width           =   2148
         End
         Begin VB.Label Label85 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Interest Warning Y/N?:"
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
            Height          =   276
            Left            =   6840
            TabIndex        =   150
            Top             =   1428
            Width           =   3300
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Billing Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   276
            Left            =   6000
            TabIndex        =   149
            Top             =   2880
            Width           =   1788
         End
         Begin VB.Label Label82 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Use Billing Cycles Y/N?:"
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
            Left            =   7920
            TabIndex        =   148
            Top             =   3360
            Width           =   2340
         End
         Begin VB.Shape Shape12 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   3228
            Left            =   120
            Top             =   2880
            Width           =   5892
         End
         Begin VB.Label Label83 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Use County Billing Y/N?:"
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
            Left            =   7920
            TabIndex        =   147
            Top             =   3840
            Width           =   2340
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Use Real/Personal Split Billing Y/N?:"
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
            Left            =   6600
            TabIndex        =   146
            Top             =   4284
            Width           =   3660
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   10368
            TabIndex        =   145
            Top             =   5208
            Width           =   300
         End
         Begin VB.Label Label87 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*CD Cash G/L:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000009&
            Height          =   276
            Left            =   3600
            TabIndex        =   144
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label Label90 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "*CD Sub G/L:"
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
            Height          =   270
            Left            =   7440
            TabIndex        =   143
            Top             =   480
            Width           =   1380
         End
      End
      Begin ImpproLib.vaImprint vaImprint5 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   159
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":915B
      End
      Begin ImpproLib.vaImprint vaImprint7 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   160
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":9177
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   8
            Left            =   7776
            TabIndex        =   161
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   4128
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   7
            Left            =   7776
            TabIndex        =   162
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   3648
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   6
            Left            =   7776
            TabIndex        =   163
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   3168
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   5
            Left            =   7776
            TabIndex        =   164
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   2688
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   4
            Left            =   7776
            TabIndex        =   165
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   2208
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   3
            Left            =   7776
            TabIndex        =   166
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   1728
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   2
            Left            =   7776
            TabIndex        =   167
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   1248
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   1
            Left            =   7776
            TabIndex        =   168
            ToolTipText     =   "Enter the hours or percentage for this distribution here."
            Top             =   768
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   1
            Left            =   1776
            TabIndex        =   169
            Top             =   768
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   2
            Left            =   1776
            TabIndex        =   170
            Top             =   1248
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   3
            Left            =   1776
            TabIndex        =   171
            Top             =   1728
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   4
            Left            =   1776
            TabIndex        =   172
            Top             =   2208
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   5
            Left            =   1776
            TabIndex        =   173
            Top             =   2688
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   6
            Left            =   1776
            TabIndex        =   174
            Top             =   3168
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   7
            Left            =   1776
            TabIndex        =   175
            Top             =   3648
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin EditLib.fpText fptxtWDAN 
            Height          =   384
            Index           =   8
            Left            =   1776
            TabIndex        =   176
            Top             =   4128
            Width           =   5628
            _Version        =   196608
            _ExtentX        =   9927
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
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Number*"
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
            Left            =   3504
            TabIndex        =   187
            Top             =   384
            Width           =   1980
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Distribution*"
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
            Left            =   7344
            TabIndex        =   186
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "1)"
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
            Left            =   1248
            TabIndex        =   185
            Top             =   864
            Width           =   300
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "2)"
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
            Left            =   1248
            TabIndex        =   184
            Top             =   1344
            Width           =   300
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "3)"
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
            Height          =   396
            Left            =   1248
            TabIndex        =   183
            Top             =   1872
            Width           =   300
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "4)"
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
            Left            =   1248
            TabIndex        =   182
            Top             =   2352
            Width           =   348
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "5)"
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
            Left            =   1248
            TabIndex        =   181
            Top             =   2832
            Width           =   300
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "6)"
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
            Left            =   1248
            TabIndex        =   180
            Top             =   3312
            Width           =   300
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "7)"
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
            Left            =   1248
            TabIndex        =   179
            Top             =   3792
            Width           =   252
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "8)"
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
            Left            =   1248
            TabIndex        =   178
            Top             =   4224
            Width           =   300
         End
         Begin VB.Label lblHrSal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hourly/$  Salary/%"
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
            Left            =   7344
            TabIndex        =   177
            Top             =   480
            Width           =   2220
         End
      End
      Begin ImpproLib.vaImprint vaImprint8 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   188
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmVATaxSystemSetup.frx":9193
         Begin LpLib.fpCombo fpcombo401K 
            Height          =   405
            Left            =   5370
            TabIndex        =   189
            ToolTipText     =   "Select Y if this employee is to be included in a 401K plan or N if otherwise."
            Top             =   4170
            Width           =   840
            _Version        =   196608
            _ExtentX        =   1482
            _ExtentY        =   714
            Text            =   ""
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
            MaxDrop         =   2
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":91AF
         End
         Begin LpLib.fpCombo fpcomboESC 
            Height          =   405
            Left            =   9030
            TabIndex        =   190
            ToolTipText     =   "Enter a :Y"" to exclude this employee on the ESC reports."
            Top             =   4170
            Width           =   840
            _Version        =   196608
            _ExtentX        =   1482
            _ExtentY        =   714
            Text            =   ""
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
            MaxDrop         =   2
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":95F6
         End
         Begin LpLib.fpCombo fpcomboLT 
            Height          =   405
            Left            =   2070
            TabIndex        =   191
            ToolTipText     =   "Select a leave table entry for this employee."
            Top             =   4170
            Width           =   840
            _Version        =   196608
            _ExtentX        =   1482
            _ExtentY        =   714
            Text            =   ""
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
            MaxDrop         =   2
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
            ColDesigner     =   "frmVATaxSystemSetup.frx":9A3D
         End
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   2
            Left            =   3264
            TabIndex        =   192
            ToolTipText     =   "Enter or adjust this employee's sick leave earned here."
            Top             =   1440
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtBal 
            CausesValidation=   0   'False
            Height          =   348
            Index           =   5
            Left            =   7488
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   3168
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtBal 
            Height          =   348
            Index           =   4
            Left            =   7488
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   2592
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtBal 
            Height          =   348
            Index           =   3
            Left            =   7488
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   2016
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            ControlType     =   1
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   8
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
         Begin EditLib.fpText fptxtBal 
            Height          =   348
            Index           =   2
            Left            =   7488
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtBal 
            Height          =   348
            Index           =   1
            Left            =   7488
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   816
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtUsed 
            Height          =   348
            Index           =   5
            Left            =   5472
            TabIndex        =   198
            ToolTipText     =   "Enter or adjust this employee's holiday time used here."
            Top             =   3168
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtUsed 
            Height          =   348
            Index           =   4
            Left            =   5472
            TabIndex        =   199
            ToolTipText     =   "Enter or adjust this employee's personal time used here."
            Top             =   2592
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtUsed 
            Height          =   348
            Index           =   1
            Left            =   5472
            TabIndex        =   200
            ToolTipText     =   "Enter or adjust the employee's vacation time used here."
            Top             =   816
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   5
            Left            =   3264
            TabIndex        =   201
            ToolTipText     =   "Enter or adjust this employee's holiday time earned here."
            Top             =   3168
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   4
            Left            =   3264
            TabIndex        =   202
            ToolTipText     =   "Enter or adjust this employee's personal time earned here."
            Top             =   2592
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   1
            Left            =   3264
            TabIndex        =   203
            ToolTipText     =   "Enter or adjust this employee's vacation time earned here."
            Top             =   864
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   3
            Left            =   3264
            TabIndex        =   204
            ToolTipText     =   "Enter or adjust this employee's comp time earned here."
            Top             =   2016
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtUsed 
            Height          =   348
            Index           =   2
            Left            =   5472
            TabIndex        =   205
            ToolTipText     =   "Enter or adjust the employee's sick leave used here."
            Top             =   1440
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtUsed 
            Height          =   348
            Index           =   3
            Left            =   5472
            TabIndex        =   206
            ToolTipText     =   "Enter or adjust this employee's comp time used here."
            Top             =   2016
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   614
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
            CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
            MaxLength       =   10
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
         Begin VB.Shape Shape10 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Vacation"
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
            Height          =   252
            Left            =   1440
            TabIndex        =   217
            Top             =   960
            Width           =   972
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Earned"
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
            Left            =   3648
            TabIndex        =   216
            Top             =   432
            Width           =   780
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Used"
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
            Left            =   5952
            TabIndex        =   215
            Top             =   384
            Width           =   540
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   7824
            TabIndex        =   214
            Top             =   336
            Width           =   876
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "Sick Leave"
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
            Height          =   252
            Left            =   1440
            TabIndex        =   213
            Top             =   1536
            Width           =   1164
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Comp Time"
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
            Left            =   1440
            TabIndex        =   212
            Top             =   2112
            Width           =   1308
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
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
            Left            =   1440
            TabIndex        =   211
            Top             =   2688
            Width           =   972
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Holiday"
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
            Left            =   1440
            TabIndex        =   210
            Top             =   3264
            Width           =   924
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Table:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   384
            Left            =   528
            TabIndex        =   209
            Top             =   4272
            Width           =   1500
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "Exclude on ESC (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   384
            Left            =   6528
            TabIndex        =   208
            Top             =   4272
            Width           =   2412
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   192
            X2              =   10128
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "401K Matching?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   384
            Left            =   3216
            TabIndex        =   207
            Top             =   4272
            Width           =   1884
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   978
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxSystemSetup.frx":9E84
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   636
      Left            =   8550
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2376
      _Version        =   131072
      _ExtentX        =   4191
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxSystemSetup.frx":A063
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNextTab 
      Height          =   636
      Left            =   6030
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2388
      _Version        =   131072
      _ExtentX        =   4212
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxSystemSetup.frx":A240
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   636
      Left            =   3522
      TabIndex        =   221
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2400
      _Version        =   131072
      _ExtentX        =   4233
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmVATaxSystemSetup.frx":A420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Required Fields = *"
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
      Height          =   660
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   228
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax System Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3144
      TabIndex        =   19
      Top             =   396
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   120
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxSystemSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim PenIdx As Integer
  Dim TempName As String
  Dim TempADD1 As String
  Dim TempADD2 As String
  Dim TempCity As String
  Dim TempTownState As String
  Dim TempZip As String
  Dim TempTaxSt As String
  Dim TempCurrYrRInt As Double
  Dim TempCurrYrPInt As Double
  Dim TempPastYrInt As Double
  Dim TempPenPct As Double
  Dim StrEmpty As Boolean
  Dim TempTaxForm As Integer
  Dim TempMinTxOpt As String
  Dim TempMinTxPct As Double
  Dim TempAcctgMethod As String
  Dim TempDisRPct As String
  Dim TempDisPPct As String
  Dim TempLateForm As String
  Dim TempCntrlDepYN As String * 1
  Dim TempCDCashGL As String * 14
  Dim TempCDSubGL As String * 14
  Dim TempPriorYrMltRevYN As String * 1
  Dim TempOverPayGLNum As String * 14
  Dim TempPenPrncTaxYN As String * 1
  Dim TempPenIntYN As String * 1
  Dim TempPenAdvYN As String * 1
  Dim TempPenLateLstYN As String * 1
  Dim TempPenPenaltyYN As String * 1
  Dim TempPenOpt1YN As String * 1
  Dim TempPenOpt2YN As String * 1
  Dim TempPenOpt3YN As String * 1
  Dim TempIntPrncTaxYN As String * 1
  Dim TempIntIntYN As String * 1
  Dim TempIntAdvYN As String * 1
  Dim TempPIntIntYN As String * 1
  Dim TempPIntAdvYN As String * 1
  Dim TempIntLateLstYN As String * 1
  Dim TempIntPenaltyYN As String * 1
  Dim TempIntOpt1YN As String * 1
  Dim TempIntOpt2YN As String * 1
  Dim TempIntOpt3YN As String * 1
  Dim TempOptRev1 As String * 35
  Dim TempOptRev2 As String * 35
  Dim TempOptRev3 As String * 35
  Dim TempPOptRev1 As String * 35
  Dim TempPOptRev2 As String * 35
  Dim TempPOptRev3 As String * 35
  Dim TempIntPOpt2YN As String * 1
  Dim TempIntPOpt3YN As String * 1
  Dim TempDisStopDate As Integer
  Dim TempOptSrchCust As String
  Dim TempOptSrchProp As String
  Dim TempOptSrchPers As String
  Dim TempWarnInt As String * 1
  Dim TempRTaxYear As Integer
  Dim TempPTaxYear As Integer
  Dim TempUseCyclesYN As String
  Dim TempPPTRAYN As String
  Dim TempUseCountyYN As String
  Dim TempRealPersSplit As String
  Dim TempSnrCtzAmt As Double
  Dim TempMaxVehVal As Double
  Dim TempMinVehVal As Double
  Dim TempMultiYear As Integer
  Dim SaveFlag As Boolean
  Dim TSListIdx As Integer
  Dim Fund As Integer, Dept As Integer, Detail As Integer
  Public RevForm As Form
  Public RealPctForm As Form
  Public PersPctForm As Form
  
Private Sub cmdAddTownship_Click()
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Integer
  Dim ThisName$
  
  On Error GoTo ERRORSTUFF
  ThisName$ = QPTrim$(fptxtTownShipName.Text)
  If ThisName$ = "" Then
    Call TaxMsg(900, "Please enter a township name.")
    fptxtTownShipName.SetFocus
    Exit Sub
  End If
  
  If InStr(cmdAddTownship.Text, "Edit") Then
    fpListTownships.Row = TSListIdx
    OpenTownshipFile TSHandle, TSCnt
    Get TSHandle, TSListIdx + 1, TSRec
    TSRec.TownShip = ThisName$
    Put TSHandle, TSListIdx + 1, TSRec
  Else
    cmdAddTownship.Text = "Add Township"
    OpenTownshipFile TSHandle, TSCnt
    TSRec.TownShip = ThisName$
    TSCnt = TSCnt + 1
    Put TSHandle, TSCnt, TSRec
  End If
      
  fpListTownships.Clear
  For x = 1 To TSCnt
    Get TSHandle, x, TSRec
    fpListTownships.AddItem QPTrim$(TSRec.TownShip)
  Next x
    
  Close TSHandle
  TSListIdx = -1
  fptxtTownShipName.Text = ""
  
  cmdAddTownship.Text = "Add Township"
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "cmdAddTownship_Click", Erl)
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

Private Sub cmdCounty_Click()
  frmVATaxCountySetup.Show vbModal
End Sub

Private Sub cmdCycle_Click()
  frmVATaxCycleSetup.Show vbModal
End Sub

Private Sub cmdExit_Click()
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  On Error GoTo ERRORSTUFF
  FileName = "C:\CPWork\exitsetup.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  Unload frmVATaxGLList
  If Exist("TaxSetup.Dat") Then
    If Check4Changes = True Then
      Exit Sub
    End If
  Else
    frmVATaxMsgWOpts.Label1.Caption = "You are exiting without saving any data. If you wish to continue exiting without saving then press F10. If you wish to return to the screen to save this data then press ESC and you will be returned to the screen where you can press the save button to save your data."
    frmVATaxMsgWOpts.Label1.Top = 600
    frmVATaxMsgWOpts.cmdCont.Text = "F10 OK To Exit"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC Abort Exit"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      GoTo ExitWOSaving
    Else
      Unload frmVATaxMsgWOpts
      vaTabPro1.ActiveTab = 0
      fptxtNameOfTaxAuth.SetFocus
      Exit Sub
    End If
  End If

  Call LogSaves
  
ExitWOSaving:
  If Exist("C:\CPWork\lateltr.dat") Then
    KillFile "C:\CPWork\lateltr.dat"
'    frmVATaxBillingMenu.Show
'    DoEvents
'    Unload RevForm
'    DoEvents
'    Unload RealPctForm
'    DoEvents
'    Unload PersPctForm
'    DoEvents
'    Unload Me
    frmVATaxBillingMenu.Show
    DoEvents
    Me.Hide
  Else
'    frmVATaxBillSetUpMenu.Show
'    DoEvents
'    Unload RevForm
'    DoEvents
'    Unload RealPctForm
'    DoEvents
'    Unload PersPctForm
'    DoEvents
'    Unload Me
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Me.Hide
  End If
  
  MainLog ("User closed frmVATaxSystemSetup.")
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "cmdExit_Click", Erl)
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

Private Sub cmdGLList_Click()
  frmVATaxGLList.Show ' vbModal
End Sub

Private Sub cmdLateBill_Click()
  Dim x As Integer
  If fpcmbLateFormat.Text = "1) SELF EDIT #1" Then
    frmVATaxLateNoticeLtr.Show
    DoEvents
  End If
End Sub

Private Sub cmdNextTab_Click()
  If vaTabPro1.ActiveTab = 0 Then
    vaTabPro1.ActiveTab = 1
    fpcmbCentDepYN.SetFocus
  ElseIf vaTabPro1.ActiveTab = 1 Then
    vaTabPro1.ActiveTab = 0
    fptxtNameOfTaxAuth.SetFocus
  End If
End Sub

Private Sub cmdPersTaxSetup_Click()
  PersPctForm.Show vbModal
End Sub

Private Sub cmdRealClass_Click()
  frmVATaxRealClassSetup.Show vbModal
End Sub

Private Sub cmdRealPctSetup_Click()
  RealPctForm.Show vbModal
End Sub

Private Sub cmdRevSetUp_Click()
  RevForm.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer, y As Integer
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim ThisRev As Integer
  Dim ThisDesc$
  Dim PenCnt As Integer
  Dim EditFlag As Boolean
  Dim ThisPen As Integer
  Dim Thisx As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  
  On Error GoTo ERRORSTUFF
  EditFlag = False
  SaveFlag = True
  If CheckPersPayOrder = False Then Exit Sub
  If fpcmbAcctMeth.Text <> "NONE" Then
    If QPTrim$(fptxtOverPayGL.Text) = "" Then
      Call TaxMsg(800, "Please enter a valid General Ledger number in the 'Overpayment GL Number' field.")
      vaTabPro1.ActiveTab = 1
      fptxtOverPayGL.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtOverPayGL.Text)) = False Then
    frmVATaxMsgWOpts.Label1.Caption = "The Overpayment GL number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"

    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      MainLog ("Warning: User issued warning that the overpayment GL number " + QPTrim$(fptxtOverPayGL.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmVATaxMsgWOpts
      Close
      vaTabPro1.ActiveTab = 1
      fptxtOverPayGL.SetFocus
      Exit Sub
    End If
  End If
    
  If QPTrim$(fptxtNameOfTaxAuth.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'Name of Taxing Authority'."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtNameOfTaxAuth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdd1.Text) = "" And QPTrim$(fptxtAdd2.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter an 'Address'."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'City' name."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtState.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'State' abbreviation for the state in which the town is located."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(ReplaceString(fptxtZip.Text, "-", "")) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'Zip Code' for the town."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbStateOfTax.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'State' abbreviation for the state in which the taxes will be paid."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbStateOfTax.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbAcctMeth.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter an 'Accounting Method'."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbAcctMeth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbTaxBillFormat.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'Tax Bill Format'."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbTaxBillFormat.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbLateFormat.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "Please enter a 'Late Bill Format'."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbLateFormat.SetFocus
    Exit Sub
  End If
  
  PenCnt = 0
  For x = 5 To 7
    RevForm.vaSpread1.Row = x
    RevForm.vaSpread1.Col = 3
    If RevForm.vaSpread1.Text = "1" Then
      ThisPen = ThisPen + 1
      Thisx = x
    End If
  Next x
  
  If ThisPen > 1 Then
    Call TaxMsg(600, "Only one optional revenue can be earmarked to be used as the penalty revenue. Please review the optional revenues selected to be used for penalty revenues and allow only one selection.")
    Close
    vaTabPro1.ActiveTab = 1
    RevForm.vaSpread1.SetActiveCell 3, Thisx
    Exit Sub
  End If
  
'  If fpcmbRPSplitYN.Enabled = False Then
  If fpcmbStateOfTax.Text = "NC" Then
    If Mid(fpcmbRPSplitYN.Text, 1, 1) = "Y" Then
      If TaxMsgWOpts(700, "The 'Use Real/Personal Split Billing Y/N?' field is set to 'Yes' but this feature is disabled. If you continue then this setting will be automatically changed to 'No'. Press F10 to continue or press ESC to stop the save procedure safely.", "F10 Continue", "ESC Stop Save") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        fpcmbCentDepYN.SetFocus
        fpcmbRPSplitYN.Enabled = False
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
        fpcmbRPSplitYN.Text = "No"
      End If
    End If
  End If
    
  If fpcmbCentDepYN.Text = "Yes" Then
    If QPTrim$(fptxtCentCash.Text) = "" Then
      Call TaxMsg(800, "You have elected to use Central Depository but no cash G/L number has been entered. Please supply a Central Depository cash G/L number before continuing.")
      vaTabPro1.ActiveTab = 1
      If fptxtCentCash.Enabled = True Then
        fptxtCentCash.SetFocus
      Else
        fpcmbCentDepYN.SetFocus
      End If
      Exit Sub
    End If
    If QPTrim$(fptxtCentSub.Text) = "" Then
      Call TaxMsg(800, "You have elected to use Central Depository but no sub G/L number has been entered. Please supply a Central Depository sub G/L number before continuing.")
      vaTabPro1.ActiveTab = 1
      If fptxtCentSub.Enabled = True Then
        fptxtCentSub.SetFocus
      Else
        fpcmbCentDepYN.SetFocus
      End If
      Exit Sub
    End If
  ElseIf fpcmbCentDepYN.Text = "No" Then
    If QPTrim$(fptxtCentCash.Text) <> "" Then
      If TaxMsgWOpts(700, "You have elected against using Central Depository. However, the Central Depository cash G/L field contains a value. This value will be deleted if you continue to save. Press F10 to continue saving. Otherwise, press ESC to abort save.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        If fptxtCentCash.Enabled = True Then
          fptxtCentCash.SetFocus
        Else
          fpcmbCentDepYN.SetFocus
        End If
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
    If QPTrim$(fptxtCentSub.Text) <> "" Then
      If TaxMsgWOpts(700, "You have elected against using Central Depository. However, the Central Depository sub G/L field contains a value. This value will be deleted if you continue to save. Press F10 to continue saving. Otherwise, press ESC to abort save.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        If fptxtCentSub.Enabled = True Then
          fptxtCentSub.SetFocus
        Else
          fpcmbCentDepYN.SetFocus
        End If
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
  End If
  
  If Exist(TaxSetupName) Then
    EditFlag = True
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
'    GoSub Check4Penalty
    If TaxMasterRec.LateForm <> CInt(Mid(fpcmbLateFormat.Text, 1, 1)) Then
      If Exist("TXLLPRN.DAT") Then
        If TaxMsgWOpts(700, "There is a late notice letter file saved. It is recommended that this file be deleted before changing to a different late notice letter format. To delete this file press F10. Otherwise press ESC to continue without deleting.", "F10 Delete", "ESC Don't Delete") = "abort" Then
          Unload frmVATaxMsgWOpts
          MainLog ("User warned that before changing the late notice format that they allow the program to delete the existing late letter printing records (TXLLPRN.DAT). The user elected NOT to delete.")
        Else
          Unload frmVATaxMsgWOpts
          KillFile ("TXLLPRN.DAT")
        End If
      End If
    End If
  Else
    OpenTaxSetUpFile TMHandle
  End If
  
  TaxMasterRec.Name = QPTrim$(fptxtNameOfTaxAuth.Text)
  TaxMasterRec.Add1 = QPTrim$(fptxtAdd1.Text)
  TaxMasterRec.Add2 = QPTrim$(fptxtAdd2.Text)
  TaxMasterRec.City = QPTrim$(fptxtCity.Text)
  TaxMasterRec.TownState = QPTrim$(fptxtState.Text)
  TaxMasterRec.Zip = QPTrim$(fptxtZip.Text)
  TaxMasterRec.TaxSt = fpcmbStateOfTax.Text
'  TaxMasterRec.CurrYrInt = fptxtCurrYrRIntRate.Value
  If Not IsNumeric(fptxtCurrRYear.Text) Then
    TaxMasterRec.RTaxYear = 0
  Else
    TaxMasterRec.RTaxYear = CInt(fptxtCurrRYear.Text)
  End If
  
  If Not IsNumeric(fptxtCurrPYear.Text) Then
    TaxMasterRec.PTaxYear = 0
  Else
    TaxMasterRec.PTaxYear = CInt(fptxtCurrPYear.Text)
  End If
  
  TaxMasterRec.PastYrInt = fptxtPastYearIntRate.Value
  TaxMasterRec.PenPct = 0 'fptxtPenaltyRate.Value
  
  Select Case fpcmbTaxBillFormat.Text
    Case "STANDARD"
      TaxMasterRec.TaxForm = 30000
'    Case "MULTI-PART"
'      TaxMasterRec.TaxForm = 21837
'    Case "POSTCARD"
'      TaxMasterRec.TaxForm = 20304
    Case "LASER"
      TaxMasterRec.TaxForm = 16716
    Case "EXPORT REAL"
      TaxMasterRec.TaxForm = 20000
    Case "EXPORT PERSONAL"
      TaxMasterRec.TaxForm = 20001
    Case "LASER ITEMIZED"
      TaxMasterRec.TaxForm = 20002
    Case "MDLTWN"
      TaxMasterRec.TaxForm = 20003
    Case "CDRBLUFF"
      TaxMasterRec.TaxForm = 20004
    Case Else
      TaxMasterRec.TaxForm = 0
  End Select
  
  TaxMasterRec.LateForm = CInt(Mid(fpcmbLateFormat.Text, 1, 1))
  TaxMasterRec.DisRPct = fptxtDiscRPct.Value
  TaxMasterRec.DisPPct = fptxtDiscPPct.Value
  TaxMasterRec.AcctgMethod = Mid(QPTrim$(fpcmbAcctMeth.Text), 1, 1)
  TaxMasterRec.OptSrchCust = QPTrim$(fptxtCustOptSrch.Text)
  TaxMasterRec.OptSrchProp = QPTrim$(fptxtPropOptSrch.Text)
  TaxMasterRec.OptSrchPers = QPTrim$(fptxtPersOptSrch.Text)
  TaxMasterRec.MinBill = fptxtMinTaxAmt.Value
  
  If InStr(fpcmbMinOptions.Text, "0") Then
    TaxMasterRec.MinTxOpt = 0
  ElseIf InStr(fpcmbMinOptions.Text, "1") Then
    TaxMasterRec.MinTxOpt = 1
  ElseIf InStr(fpcmbMinOptions.Text, "2") Then
    TaxMasterRec.MinTxOpt = 2
  Else
    TaxMasterRec.MinTxOpt = 0
  End If
  
  TaxMasterRec.CntrlDepYN = fpcmbCentDepYN.Text
  If fptxtCentCash.Enabled = True Then
    TaxMasterRec.CDCashGL = QPTrim$(fptxtCentCash.Text)
  Else
    TaxMasterRec.CDCashGL = ""
  End If
  If fptxtCentSub.Enabled = True Then
    TaxMasterRec.CDSubGL = QPTrim$(fptxtCentSub.Text)
  Else
    TaxMasterRec.CDSubGL = ""
  End If
  
  TaxMasterRec.OverPayGLNum = QPTrim$(fptxtOverPayGL.Text)
  TaxMasterRec.PriorYrMltRevYN = "N" 'fpcmbMultRevYN.Text
  TaxMasterRec.WarnInt = fpcmbNoInterYN.Text
  TaxMasterRec.MultiYear = fpcmbMultiYear.Text
  TaxMasterRec.PPTRADisc = CDbl(fpDSPPTRADisc.Value)
  For x = 1 To 8
    RevForm.vaSpread1.Row = x
    RevForm.vaSpread1.Col = 1
    If x = 6 Then
      If QPTrim$(RevForm.vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev1) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread1.Text) + " as an optional revenue on row " + CStr(RevForm.vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread1.SetFocus
            RevForm.vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        RevForm.vaSpread1.Col = 3
        If QPTrim$(RevForm.vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #1 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          RevForm.Show vbModal
          RevForm.vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        RevForm.vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev1) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev1) <> QPTrim$(RevForm.vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #1. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Close
             RevForm.Show vbModal
             RevForm.vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) <> QPTrim$(TaxMasterRec.OptRev1) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #1 from " + QPTrim$(TaxMasterRec.OptRev1) + " to " + QPTrim$(RevForm.vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev1 = QPTrim$(RevForm.vaSpread1.Text)
    End If
    If x = 7 Then
      If QPTrim$(RevForm.vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev2) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread1.Text) + " as an optional revenue on row " + CStr(RevForm.vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        RevForm.vaSpread1.Col = 3
        If QPTrim$(RevForm.vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #2 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          RevForm.Show vbModal
          RevForm.vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        RevForm.vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev2) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev2) <> QPTrim$(RevForm.vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #2. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmVATaxMsgWOpts
             Close
             RevForm.Show vbModal
             RevForm.vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) <> QPTrim$(TaxMasterRec.OptRev2) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #2 from " + QPTrim$(TaxMasterRec.OptRev2) + " to " + QPTrim$(RevForm.vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev2 = QPTrim$(RevForm.vaSpread1.Text)
    End If
    If x = 8 Then
      If QPTrim$(RevForm.vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev3) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread1.Text) + " as an optional revenue on row " + CStr(RevForm.vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        RevForm.vaSpread1.Col = 3
        If QPTrim$(RevForm.vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #3 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          Unload frmVATaxMsgWOpts
          RevForm.Show vbModal
          RevForm.vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        RevForm.vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev3) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev3) <> QPTrim$(RevForm.vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #3. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmVATaxMsgWOpts
             Close
             RevForm.Show vbModal
             RevForm.vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) <> QPTrim$(TaxMasterRec.OptRev3) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #3 from " + QPTrim$(TaxMasterRec.OptRev3) + " to " + QPTrim$(RevForm.vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev3 = QPTrim$(RevForm.vaSpread1.Text)
    End If
    
    RevForm.vaSpread1.Col = 2
    Select Case x
      Case 1
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntIntYN = "Y"
        Else
          TaxMasterRec.IntIntYN = "N"
        End If
      Case 2
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntAdvYN = "Y"
        Else
          TaxMasterRec.IntAdvYN = "N"
        End If
      Case 3
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntLateLstYN = "Y"
        Else
          TaxMasterRec.IntLateLstYN = "N"
        End If
      Case 4
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntPenaltyYN = "Y"
        Else
          TaxMasterRec.IntPenaltyYN = "N"
        End If
      Case 5
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntPrncTaxYN = "Y"
        Else
          TaxMasterRec.IntPrncTaxYN = "N"
        End If
      Case 6
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt1YN = "Y"
        Else
          TaxMasterRec.IntOpt1YN = "N"
        End If
      Case 7
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt2YN = "Y"
        Else
          TaxMasterRec.IntOpt2YN = "N"
        End If
      Case 8
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt3YN = "Y"
        Else
          TaxMasterRec.IntOpt3YN = "N"
        End If
    End Select
    
'    TaxMasterRec.CurrRYrInt(1) = RealPctForm.fpDblYearInt(0).Value
'    TaxMasterRec.CurrRYrInt(2) = RealPctForm.fpDblYearInt(1).Value
'    TaxMasterRec.CurrRYrInt(3) = RealPctForm.fpDblYearInt(2).Value
'    TaxMasterRec.CurrRYrInt(4) = RealPctForm.fpDblYearInt(3).Value
'    TaxMasterRec.CurrRYrInt(5) = RealPctForm.fpDblYearInt(4).Value
'    TaxMasterRec.CurrRYrIntInUse = CDbl(fptxtCurrYrRIntRate.Value)
'
'    TaxMasterRec.CurrPYrInt(1) = PersPctForm.fpDblYearInt(0).Value
'    TaxMasterRec.CurrPYrInt(2) = PersPctForm.fpDblYearInt(1).Value
'    TaxMasterRec.CurrPYrInt(3) = PersPctForm.fpDblYearInt(2).Value
'    TaxMasterRec.CurrPYrInt(4) = PersPctForm.fpDblYearInt(3).Value
'    TaxMasterRec.CurrPYrInt(5) = PersPctForm.fpDblYearInt(4).Value
'    TaxMasterRec.CurrPYrIntInUse = CDbl(fptxtCurrYrPIntRate.Value)
   
'    RevForm.vaSpread1.Col = 3
'    Select Case x
'      Case 5
'        If RevForm.vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            RevForm.vaSpread1.Col = 1
'            frmVATaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmVATaxMsgWOpts.Label1.Top = 700
'            frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmVATaxMsgWOpts.Show vbModal
'            If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmVATaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmVATaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 5
'        End If
'      Case 6
'        If RevForm.vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            frmVATaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmVATaxMsgWOpts.Label1.Top = 700
'            frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmVATaxMsgWOpts.Show vbModal
'            If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmVATaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmVATaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 6
'        End If
'      Case 7
'        If RevForm.vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            frmVATaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmVATaxMsgWOpts.Label1.Top = 700
'            frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmVATaxMsgWOpts.Show vbModal
'            If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmVATaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmVATaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 7
'        End If
'    End Select
    RevForm.vaSpread1.Col = 4
    Select Case x
      Case 1
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenIntYN = "Y"
        Else
          TaxMasterRec.PenIntYN = "N"
        End If
      Case 2
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenAdvYN = "Y"
        Else
          TaxMasterRec.PenAdvYN = "N"
        End If
      Case 3
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenLateLstYN = "Y"
        Else
          TaxMasterRec.PenLateLstYN = "N"
        End If
      Case 4
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenPenaltyYN = "Y"
        Else
          TaxMasterRec.PenPenaltyYN = "N"
        End If
      Case 5
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenPrncTaxYN = "Y"
        Else
          TaxMasterRec.PenPrncTaxYN = "N"
        End If
      Case 6
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt1YN = "Y"
        Else
          TaxMasterRec.PenOpt1YN = "N"
        End If
      Case 7
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt2YN = "Y"
        Else
          TaxMasterRec.PenOpt2YN = "N"
        End If
      Case 8
        If RevForm.vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt3YN = "Y"
        Else
          TaxMasterRec.PenOpt3YN = "N"
        End If
    End Select
  Next x
  
  '-----------------------------------------------------------------
  
  For x = 1 To 10
    RevForm.vaSpread2.Row = x
    RevForm.vaSpread2.Col = 1
    If x = 8 Then
      If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev1) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread2.Text) + " as a personal optional revenue on row " + CStr(RevForm.vaSpread2.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread2.SetFocus
            RevForm.vaSpread2.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.POptRev1) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev1) <> QPTrim$(RevForm.vaSpread2.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of personal optional revenue #1. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Close
             RevForm.Show vbModal
             RevForm.vaSpread2.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.POptRev1) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread2.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of personal optional revenue #1 from " + QPTrim$(TaxMasterRec.POptRev1) + " to " + QPTrim$(RevForm.vaSpread2.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.POptRev1 = QPTrim$(RevForm.vaSpread2.Text)
    End If
    If x = 9 Then
      If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev2) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread2.Text) + " as a personal optional revenue on row " + CStr(RevForm.vaSpread2.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread2.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.POptRev2) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev2) <> QPTrim$(RevForm.vaSpread2.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of personal optional revenue #2. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmVATaxMsgWOpts
             Close
             RevForm.Show vbModal
             RevForm.vaSpread2.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.POptRev2) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread2.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of personal optional revenue #2 from " + QPTrim$(TaxMasterRec.POptRev2) + " to " + QPTrim$(RevForm.vaSpread2.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.POptRev2 = QPTrim$(RevForm.vaSpread2.Text)
    End If
    If x = 10 Then
      If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev3) = "" Then
          frmVATaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(RevForm.vaSpread2.Text) + " as a personal optional revenue on row " + CStr(RevForm.vaSpread2.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmVATaxMsgWOpts.Label1.Top = 700
          frmVATaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmVATaxMsgWOpts.Show vbModal
          If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmVATaxMsgWOpts
            Close
            RevForm.Show vbModal
            RevForm.vaSpread2.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.POptRev3) <> "" Then
        If QPTrim$(TaxMasterRec.POptRev3) <> QPTrim$(RevForm.vaSpread2.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of personal optional revenue #3. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmVATaxMsgWOpts
             Close
             RevForm.Show vbModal
             RevForm.vaSpread2.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.POptRev3) Then
                     TblRec.Desc = QPTrim$(RevForm.vaSpread2.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of personal optional revenue #3 from " + QPTrim$(TaxMasterRec.POptRev3) + " to " + QPTrim$(RevForm.vaSpread2.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.POptRev3 = QPTrim$(RevForm.vaSpread2.Text)
    End If
  
    RevForm.vaSpread2.Col = 2
    Select Case x
      Case 1
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPersYN = "Y"
        Else
          TaxMasterRec.IntPersYN = "N"
        End If
      Case 2
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntMTYN = "Y"
        Else
          TaxMasterRec.IntMTYN = "N"
        End If
      Case 3
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntMCYN = "Y"
        Else
          TaxMasterRec.IntMCYN = "N"
        End If
      Case 4
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntFEYN = "Y"
        Else
          TaxMasterRec.IntFEYN = "N"
        End If
      Case 5
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntMHYN = "Y"
        Else
          TaxMasterRec.IntMHYN = "N"
        End If
      Case 6
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPIntYN = "Y"
        Else
          TaxMasterRec.IntPIntYN = "N"
        End If
      Case 7
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPPenYN = "Y"
        Else
          TaxMasterRec.IntPPenYN = "N"
        End If
      Case 8
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPOpt1YN = "Y"
        Else
          TaxMasterRec.IntPOpt1YN = "N"
        End If
      Case 9
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPOpt2YN = "Y"
        Else
          TaxMasterRec.IntPOpt2YN = "N"
        End If
      Case 10
        If RevForm.vaSpread2.Text = "1" Then
          TaxMasterRec.IntPOpt3YN = "Y"
        Else
          TaxMasterRec.IntPOpt3YN = "N"
        End If
    End Select
    
    RevForm.vaSpread2.Col = 3
    Select Case x
      Case 1
        TaxMasterRec.PersPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 2
        TaxMasterRec.MTPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 3
        TaxMasterRec.MCPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 4
        TaxMasterRec.FEPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 5
        TaxMasterRec.MHPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 6
        TaxMasterRec.PIntPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 7
        TaxMasterRec.PPenPayOrder = CInt(RevForm.vaSpread2.Text)
      Case 8
        TaxMasterRec.POpt1PayOrder = CInt(RevForm.vaSpread2.Text)
      Case 9
        TaxMasterRec.POpt2PayOrder = CInt(RevForm.vaSpread2.Text)
      Case 10
        TaxMasterRec.POpt3PayOrder = CInt(RevForm.vaSpread2.Text)
    End Select
  Next x
  
  '-----------------------------------------------------------------
  If EditFlag = False Then
    TaxMasterRec.DiscRXDate = 0
    TaxMasterRec.DiscPXDate = 0
  End If
  TaxMasterRec.PPTRAYN = fpcmbPPTRAYN.Text
  TaxMasterRec.UseCyclesYN = Mid(fpcmbCyclesYN.Text, 1, 1)
  TaxMasterRec.UseCountyYN = Mid(fpcmbCountyYN.Text, 1, 1)
  TaxMasterRec.RealPersSplit = Mid(fpcmbRPSplitYN.Text, 1, 1)
  TaxMasterRec.MaxVehTaxVal = CDbl(fpCurrMaxVehAmt.Value)
  TaxMasterRec.MinVehTaxVal = CDbl(fpCurrMinVehAmt.Value)
  If fptxtLawChngDate.Text = "N/A" Then
    TaxMasterRec.LawChngDate = 0
  Else
    TaxMasterRec.LawChngDate = Date2Num(fptxtLawChngDate.Text)
  End If
  TaxMasterRec.CurrRYrInt(1) = RealPctForm.fpDblYearInt(0).Value
  TaxMasterRec.CurrRYrInt(2) = RealPctForm.fpDblYearInt(1).Value
  TaxMasterRec.CurrRYrInt(3) = RealPctForm.fpDblYearInt(2).Value
  TaxMasterRec.CurrRYrInt(4) = RealPctForm.fpDblYearInt(3).Value
  TaxMasterRec.CurrRYrInt(5) = RealPctForm.fpDblYearInt(4).Value
  TaxMasterRec.CurrRYrIntInUse = CDbl(fptxtCurrYrRIntRate.Value)
    
  TaxMasterRec.CurrPYrInt(1) = PersPctForm.fpDblYearInt(0).Value
  TaxMasterRec.CurrPYrInt(2) = PersPctForm.fpDblYearInt(1).Value
  TaxMasterRec.CurrPYrInt(3) = PersPctForm.fpDblYearInt(2).Value
  TaxMasterRec.CurrPYrInt(4) = PersPctForm.fpDblYearInt(3).Value
  TaxMasterRec.CurrPYrInt(5) = PersPctForm.fpDblYearInt(4).Value
  TaxMasterRec.CurrPYrIntInUse = CDbl(fptxtCurrYrPIntRate.Value)
  Put TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Call LogSaves
  
  Unload frmVATaxGLList
  Call Savemsg(900, "Your data has been saved successfully.")
  
  If Exist("C:\CPWork\ratetbls.dat") Then
    KillFile "C:\CPWork\ratetbls.dat"
    RevForm.vaSpread1.Col = 1
    For x = 5 To 7
      RevForm.vaSpread1.Row = x
      If QPTrim$(RevForm.vaSpread1.Text) <> "" Then
        frmVATaxRateMenu.Show
        DoEvents
        Unload Me
        Exit Sub
      End If
    Next x
  End If
  
  If Exist("C:\CPWork\lateltr.dat") Then
    KillFile "C:\CPWork\lateltr.dat"
    frmVATaxBillingMenu.Show
    DoEvents
    Unload RevForm
    Unload Me
  Else
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Unload RevForm
    Unload Me
  End If
  
  Exit Sub
  
Check4Penalty:
  PenCnt = 0
  If PenIdx = 0 Then Return
  If TaxMasterRec.PenIdx = 0 Then Return
  
  RevForm.vaSpread1.Col = 2
  For y = 6 To 8
    RevForm.vaSpread1.Row = y
    If RevForm.vaSpread1.Text = "1" Then
      PenCnt = PenCnt + 1
        RevForm.vaSpread1.Col = 1
        ThisDesc$ = QPTrim$(RevForm.vaSpread1.Text)
        If ThisDesc = "" Then
          frmVATaxMsg.Label1.Caption = "You have designated revenue #" + CStr(y) + " as your penalty revenue but no description has been entered. Please enter a description for this penalty revenue."
          frmVATaxMsg.Show vbModal
          Close
          RevForm.vaSpread1.SetActiveCell 1, y
          Exit Sub
        End If
    End If
  Next y
    
'  If PenCnt > 0 And fptxtPenaltyRate.Value = 0 Then
'    frmVATaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is zero. Press F10 if you wish to continue saving anyway. Press ESC to review."
'    frmVATaxMsgWOpts.Label1.Top = 800
'    frmVATaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
'    frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
'    frmVATaxMsgWOpts.Show vbModal
'    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmVATaxMsgWOpts
'      fptxtPenaltyRate.SetFocus
'      Close
'      Exit Sub
'    ElseIf frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
'      Unload frmVATaxMsgWOpts
'    End If
'  End If
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "cmdSave", Erl)
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
  
End Sub

Private Sub cmdSave_GotFocus()
  SaveFlag = True
End Sub

Private Sub cmdTaxBill_Click()
  Dim x As Integer
  If fpcmbTaxBillFormat.Text = "LASER" Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      frmVATaxRealBillPrint.Show
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      frmVATaxPersBillPrint.Show
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf InStr(fpcmbTaxBillFormat.Text, "EXPORT") Then
    Call TaxMsg(800, "The 'EXPORT' bill format creates a file that is then forwarded to a tax billing company. There is no hard copy tax bill to display.")
    Exit Sub
  ElseIf fpcmbTaxBillFormat.Text = "STANDARD" Then
    frmVATaxStandardBill.Show
  ElseIf fpcmbTaxBillFormat.Text = "LASER ITEMIZED" Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      frmVATaxRealBillPrint.Show
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      frmVATaxPersLsrItemized.Show
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf fpcmbTaxBillFormat.Text = "MDLTWN" Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      Call PrintMdltwnReal
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      Call PrintMdltwnPers
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
  ElseIf fpcmbTaxBillFormat.Text = "CDRBLUFF" Then
    frmVATaxBillPostOpt.Show vbModal
    If frmVATaxBillPostOpt.fptxtPostType.Text = "Real" Then
      Call PrintCdrBluffReal
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Personal" Then
      Call PrintCdrBluffPers
      DoEvents
      Unload frmVATaxBillPostOpt
    ElseIf frmVATaxBillPostOpt.fptxtPostType.Text = "Exit" Then
      DoEvents
      Unload frmVATaxBillPostOpt
      Exit Sub
    End If
    
  End If
  DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  fpcmbRPSplitYN.Enabled = True
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxSystemSetup.")
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
    Call cmdNextTab_Click
    KeyCode = 0
  End If
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
    Case vbKeyF5:
      SendKeys "%N"
      Call cmdNextTab_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%G"
      Call cmdGLList_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%w"
      Call cmdTaxBill_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%B"
      Call cmdLateBill_Click
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
      Unload frmVATaxGLList
      Unload RevForm
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxSystemSetup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MinTaxString$
  Dim x As Integer
  Dim ThisCol As Integer
  Dim PenRec As PenaltyHandlingType
  Dim PHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim DateLen As Integer
  Dim ThisYr$
  Dim TempIntPersYN$
  Dim TempIntMTYN$
  Dim TempIntMCYN
  Dim TempIntFEYN
  Dim TempIntMHYN
  Dim TempPIntOpt1YN
  Dim TempPIntOpt2YN
  Dim TempIntPIntYN
  Dim TempIntPPenYN
  
  On Error GoTo ERRORSTUFF
  
  Set RevForm = New frmVATaxRevSpreadsheets
  RevForm.Show
  RevForm.Hide
  Set RealPctForm = New frmVATaxRealPctSetup
  RealPctForm.Show
  RealPctForm.Hide
  Set PersPctForm = New frmVATaxPersPctSetup
  PersPctForm.Show
  PersPctForm.Hide
  
  If Not Exist("GLACCT.IDX") Or Not Exist("GLACCT.DAT") Then
    Fund = 0
    Dept = 0
    Detail = 0
  Else
    Call GetAcctStruct(CurrCitiPath, Fund, Dept, Detail)
  End If
    
  Me.HelpContextID = hlpTaxSystemDefault
  If Exist(TaxTownships) Then
    fpListTownships.Clear
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpListTownships.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  Else
    fpListTownships.AddItem "No Townships Saved"
  End If
      
  SaveFlag = False
  PenIdx = 0
  
  If Exist("TAXSETUP.Dat") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
  End If
  
  PenIdx = TaxMasterRec.PenIdx
  vaTabPro1.ActiveTab = 0
  fptxtNameOfTaxAuth.Text = QPTrim$(TaxMasterRec.Name)
  TempName = QPTrim$(TaxMasterRec.Name)
  fptxtAdd1.Text = QPTrim$(TaxMasterRec.Add1)
  TempADD1 = QPTrim$(TaxMasterRec.Add1)
  fptxtAdd2.Text = QPTrim$(TaxMasterRec.Add2)
  TempADD2 = QPTrim$(TaxMasterRec.Add2)
  fptxtCity.Text = QPTrim$(TaxMasterRec.City)
  TempCity = QPTrim$(TaxMasterRec.City)
  fptxtState.Text = TaxMasterRec.TownState
  TempTownState = TaxMasterRec.TownState
  fptxtZip.Text = QPTrim$(TaxMasterRec.Zip)
  TempZip = QPTrim$(TaxMasterRec.Zip)
  fpcmbStateOfTax.Text = QPTrim$(TaxMasterRec.TaxSt)
  TempTaxSt = QPTrim$(TaxMasterRec.TaxSt)
  Select Case TaxMasterRec.TaxForm
'    Case 21837
'      fpcmbTaxBillFormat.Text = "MULTI-PART"
'    Case 20304
'      fpcmbTaxBillFormat.Text = "POSTCARD"
    Case 16716
      fpcmbTaxBillFormat.Text = "LASER"
    Case 30000
      fpcmbTaxBillFormat.Text = "STANDARD"
    Case 20000
      fpcmbTaxBillFormat.Text = "EXPORT REAL"
    Case 20001
      fpcmbTaxBillFormat.Text = "EXPORT PERSONAL"
    Case 20002
      fpcmbTaxBillFormat.Text = "LASER ITEMIZED"
    Case 20003
      fpcmbTaxBillFormat.Text = "MDLTWN"
    Case 20004
      fpcmbTaxBillFormat.Text = "CDRBLUFF"
    Case Else
      fpcmbTaxBillFormat.Text = "UNKNOWN"
  End Select
  TempTaxForm = TaxMasterRec.TaxForm
  
  fpcmbTaxBillFormat.AddItem "STANDARD"
  fpcmbTaxBillFormat.AddItem "LASER"
  fpcmbTaxBillFormat.AddItem "EXPORT REAL"
  fpcmbTaxBillFormat.AddItem "EXPORT PERSONAL"
  fpcmbTaxBillFormat.AddItem "LASER ITEMIZED"
  fpcmbTaxBillFormat.AddItem "MDLTWN"
  fpcmbTaxBillFormat.AddItem "CDRBLUFF"
  
  fpcmbStateOfTax.AddItem "NC"
  fpcmbStateOfTax.AddItem "VA"
  fpcmbStateOfTax.AddItem "GA"
  fpcmbStateOfTax.AddItem "SC"
  fpcmbStateOfTax.AddItem "MD"
  fpcmbStateOfTax.AddItem "TN"
  
  Select Case TaxMasterRec.AcctgMethod
    Case "N"
      fpcmbAcctMeth.Text = "NONE"
    Case "M"
      fpcmbAcctMeth.Text = "MODIFIED ACCRUAL"
    Case "A"
      fpcmbAcctMeth.Text = "ACCRUAL"
    Case "C"
      fpcmbAcctMeth.Text = "CASH"
    Case Else
      fpcmbAcctMeth.Text = "NONE"
  End Select
  TempAcctgMethod = TaxMasterRec.AcctgMethod
  fpcmbAcctMeth.AddItem "NONE"
  fpcmbAcctMeth.AddItem "MODIFIED ACCRUAL"
  fpcmbAcctMeth.AddItem "ACCRUAL"
  fpcmbAcctMeth.AddItem "CASH"
  
  MinTaxString = CStr(TaxMasterRec.MinBill)
  If InStr(MinTaxString, "E") Then
    TaxMasterRec.MinBill = 0
  End If
  fptxtMinTaxAmt = TaxMasterRec.MinBill
  TempMinTxPct = TaxMasterRec.MinBill
  
  fptxtCustOptSrch.Text = QPTrim$(TaxMasterRec.OptSrchCust)
  TempOptSrchCust = QPTrim$(TaxMasterRec.OptSrchCust)
  fptxtPropOptSrch.Text = QPTrim$(TaxMasterRec.OptSrchProp)
  TempOptSrchProp = QPTrim$(TaxMasterRec.OptSrchProp)
  fptxtPersOptSrch.Text = QPTrim$(TaxMasterRec.OptSrchPers)
  TempOptSrchPers = QPTrim$(TaxMasterRec.OptSrchPers)
  Select Case TaxMasterRec.MinTxOpt
    Case 0
      fpcmbMinOptions.Text = "(0) No special minimum tax handling."
    Case 1
      fpcmbMinOptions.Text = "(1) Charge no tax if tax bill is at or less than minimum."
    Case 2
      fpcmbMinOptions.Text = "(2) Charge minimum if tax bill is at or less than minimum."
    Case Else
      fpcmbMinOptions.Text = "(0) No special minimum tax handling."
  End Select
  
  TempMinTxOpt = TaxMasterRec.MinTxOpt
  fpcmbMinOptions.AddItem "(0) No special minimum tax handling."
  fpcmbMinOptions.AddItem "(1) Charge no tax if tax bill is at or less than minimum."
  fpcmbMinOptions.AddItem "(2) Charge minimum if tax bill is at or less than minimum."
  
  fptxtCurrYrRIntRate.Text = TaxMasterRec.CurrRYrIntInUse
  TempCurrYrRInt = TaxMasterRec.CurrRYrIntInUse
  fptxtCurrYrPIntRate.Text = TaxMasterRec.CurrPYrIntInUse
  TempCurrYrPInt = TaxMasterRec.CurrPYrIntInUse
 
  DateLen = Len(Date)
  ThisYr = Mid(Date, DateLen - 3, DateLen)
  If TaxMasterRec.RTaxYear = 0 Then
    fptxtCurrRYear.Text = ThisYr
    TempRTaxYear = 0
  Else
    fptxtCurrRYear.Text = CStr(TaxMasterRec.RTaxYear)
    TempRTaxYear = TaxMasterRec.RTaxYear
  End If
  
  If TaxMasterRec.PTaxYear = 0 Then
    fptxtCurrPYear.Text = ThisYr
    TempPTaxYear = 0
  Else
    fptxtCurrPYear.Text = CStr(TaxMasterRec.PTaxYear)
    TempPTaxYear = TaxMasterRec.PTaxYear
  End If
  
  fptxtPastYearIntRate.Text = TaxMasterRec.PastYrInt
  TempPastYrInt = TaxMasterRec.PastYrInt
  
  '-------------Tab2-------------------
  If TaxMasterRec.LateForm = 1 Then
    fpcmbLateFormat.Text = "1) SELF EDIT #1"
  ElseIf TaxMasterRec.LateForm = 0 Then
    fpcmbLateFormat.Text = "0) None Saved"
  End If
  fpcmbLateFormat.AddItem "0) None Saved"
  fpcmbLateFormat.AddItem "1) SELF EDIT #1"
  TempLateForm = TaxMasterRec.LateForm
  
'  fptxtPenaltyRate.Text = TaxMasterRec.PenPct
'  TempPenPct = TaxMasterRec.PenPct
  
  fptxtDiscRPct.Text = TaxMasterRec.DisRPct
  TempDisRPct = TaxMasterRec.DisRPct
  fptxtDiscPPct.Text = TaxMasterRec.DisPPct
  TempDisPPct = TaxMasterRec.DisPPct
  
  If TaxMasterRec.CntrlDepYN = "Y" Then
    fpcmbCentDepYN.Text = "Yes"
  Else
    fpcmbCentDepYN.Text = "No"
  End If
  TempCntrlDepYN = TaxMasterRec.CntrlDepYN
  fpcmbCentDepYN.AddItem "Yes"
  fpcmbCentDepYN.AddItem "No"
  
  fptxtCentCash.Text = QPTrim$(TaxMasterRec.CDCashGL)
  TempCDCashGL = QPTrim$(TaxMasterRec.CDCashGL)
  
  fptxtCentSub.Text = QPTrim$(TaxMasterRec.CDSubGL)
  TempCDSubGL = QPTrim$(TaxMasterRec.CDSubGL)
  
'  If TaxMasterRec.PriorYrMltRevYN = "Y" Then
'    fpcmbMultRevYN.Text = "Yes"
'  Else
'    fpcmbMultRevYN.Text = "No"
'  End If
'  TempPriorYrMltRevYN = TaxMasterRec.PriorYrMltRevYN
'  fpcmbMultRevYN.AddItem "Yes"
'  fpcmbMultRevYN.AddItem "No"
  
  If TaxMasterRec.WarnInt = "Y" Then
    fpcmbNoInterYN.Text = "Yes"
  Else
    fpcmbNoInterYN.Text = "No"
  End If
  
  fpcmbNoInterYN.AddItem "Yes"
  fpcmbNoInterYN.AddItem "No"
  
  TempWarnInt = fpcmbNoInterYN.Text
  
  For x = 1 To 12
    fpcmbMultiYear.AddItem CStr(x)
  Next x
  If TaxMasterRec.MultiYear = 0 Then TaxMasterRec.MultiYear = 1
  fpcmbMultiYear.Text = CStr(TaxMasterRec.MultiYear)
  fpDSPPTRADisc = TaxMasterRec.PPTRADisc
  fptxtOverPayGL.Text = QPTrim$(TaxMasterRec.OverPayGLNum)
  TempOverPayGLNum = QPTrim$(TaxMasterRec.OverPayGLNum)
  fpCurrMaxVehAmt = TaxMasterRec.MaxVehTaxVal
  fpCurrMinVehAmt = TaxMasterRec.MinVehTaxVal
  TempMaxVehVal = TaxMasterRec.MaxVehTaxVal
  TempMinVehVal = TaxMasterRec.MinVehTaxVal
  TempMultiYear = TaxMasterRec.MultiYear
  
  ReDim Sprd1Col1(1 To 8) As String
  ReDim Sprd1Col2(1 To 8) As Integer
  ReDim Sprd2Col1(1 To 10) As String
  ReDim Sprd2Col2(1 To 10) As Integer
  ReDim Sprd2Col3(1 To 10) As Integer
  For x = 1 To 8
    RevForm.vaSpread1.Row = x
    RevForm.vaSpread1.Col = 1
    Select Case x
      Case 1
        RevForm.vaSpread1.Text = "Default: Interest Accrued"
        RevForm.vaSpread1.Lock = True
        Sprd1Col1(1) = "Default: Interest Accrued"
      Case 2
        RevForm.vaSpread1.Text = "Default: Advertising Cost Incurred"
        RevForm.vaSpread1.Lock = True
        Sprd1Col1(2) = "Default: Advertising Cost Incurred"
      Case 3
        RevForm.vaSpread1.Text = "Default: Late Listing"
        RevForm.vaSpread1.Lock = True
        Sprd1Col1(3) = "Default: Late Listing"
      Case 4
        RevForm.vaSpread1.Text = "Default: Penalty"
        RevForm.vaSpread1.Lock = True
        Sprd1Col1(4) = "Default: Penalty"
      Case 5
        RevForm.vaSpread1.Text = "Default: Principle"
        RevForm.vaSpread1.Lock = True
        Sprd1Col1(5) = "Default: Principle"
      Case 6
        RevForm.vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev1)
        TempOptRev1 = TaxMasterRec.OptRev1
        Sprd1Col1(6) = QPTrim$(TaxMasterRec.OptRev1)
      Case 7
        RevForm.vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev2)
        TempOptRev2 = TaxMasterRec.OptRev2
        Sprd1Col1(7) = QPTrim$(TaxMasterRec.OptRev2)
      Case 8
        RevForm.vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev3)
        TempOptRev3 = TaxMasterRec.OptRev3
        Sprd1Col1(8) = QPTrim$(TaxMasterRec.OptRev3)
      Case Else
    End Select
    RevForm.vaSpread1.Col = 2
    Select Case x
      Case 1
        If TaxMasterRec.IntIntYN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(1) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(1) = 0
        End If
        TempIntIntYN = TaxMasterRec.IntIntYN
      Case 2
        If TaxMasterRec.IntAdvYN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(2) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(2) = 0
        End If
        TempIntAdvYN = TaxMasterRec.IntAdvYN
      Case 3
        If TaxMasterRec.IntLateLstYN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(3) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(3) = 0
        End If
        TempIntLateLstYN = TaxMasterRec.IntLateLstYN
      Case 4
        If TaxMasterRec.IntPenaltyYN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(4) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(4) = 0
        End If
        TempIntPenaltyYN = TaxMasterRec.IntPenaltyYN
      Case 5
        If TaxMasterRec.IntPrncTaxYN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(5) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(5) = 0
        End If
        TempIntPrncTaxYN = TaxMasterRec.IntPrncTaxYN
      Case 6
        If TaxMasterRec.IntOpt1YN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(6) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(6) = 0
        End If
        TempIntOpt1YN = TaxMasterRec.IntOpt1YN
      Case 7
        If TaxMasterRec.IntOpt2YN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(7) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(7) = 0
        End If
        TempIntOpt2YN = TaxMasterRec.IntOpt2YN
      Case 8
        If TaxMasterRec.IntOpt3YN = "Y" Then
          RevForm.vaSpread1.Value = 1
          Sprd1Col2(8) = 1
        Else
          RevForm.vaSpread1.Value = 0
          Sprd1Col2(8) = 0
        End If
        TempIntOpt3YN = TaxMasterRec.IntOpt3YN
      Case Else
    End Select
    
    RevForm.vaSpread1.Col = 3
    Select Case x
      Case 1
        RevForm.vaSpread1.Lock = True
      Case 2
        RevForm.vaSpread1.Lock = True
      Case 3
        RevForm.vaSpread1.Lock = True
      Case 4
        RevForm.vaSpread1.Lock = True
      Case 5
        RevForm.vaSpread1.Lock = True
      Case 6
        If PenIdx = 5 Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        
      Case 7
        If TaxMasterRec.PenIdx = 6 Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
      Case 8
        If TaxMasterRec.PenIdx = 7 Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
      Case Else
    End Select
    
    RevForm.vaSpread1.Col = 4
    Select Case x
      Case 1
        If TaxMasterRec.PenIntYN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenIntYN = TaxMasterRec.PenIntYN
      Case 2
        If TaxMasterRec.PenAdvYN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenAdvYN = TaxMasterRec.PenAdvYN
      Case 3
        If TaxMasterRec.PenLateLstYN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenLateLstYN = TaxMasterRec.PenLateLstYN
      Case 4
        If TaxMasterRec.PenPenaltyYN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenPenaltyYN = TaxMasterRec.PenPenaltyYN
      Case 5
        If TaxMasterRec.PenPrncTaxYN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenPrncTaxYN = TaxMasterRec.PenPrncTaxYN
      Case 6
        If TaxMasterRec.PenOpt1YN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenOpt1YN = TaxMasterRec.PenOpt1YN
      Case 7
        If TaxMasterRec.PenOpt2YN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenOpt2YN = TaxMasterRec.PenOpt2YN
      Case 8
        If TaxMasterRec.PenOpt3YN = "Y" Then
          RevForm.vaSpread1.Value = 1
        Else
          RevForm.vaSpread1.Value = 0
        End If
        TempPenOpt3YN = TaxMasterRec.PenOpt3YN
      Case Else
    End Select
  Next x
  
  For x = 1 To 10
    RevForm.vaSpread2.Row = x
    RevForm.vaSpread2.Col = 1
    Select Case x
      Case 1
        RevForm.vaSpread2.Lock = True
      Case 2
        RevForm.vaSpread2.Lock = True
      Case 3
        RevForm.vaSpread2.Lock = True
      Case 4
        RevForm.vaSpread2.Lock = True
      Case 5
        RevForm.vaSpread2.Lock = True
      Case 6
        RevForm.vaSpread2.Lock = True
      Case 7
        RevForm.vaSpread2.Lock = True
    End Select
    Select Case x
      Case 1
        RevForm.vaSpread2.Text = "Personal Tax"
        Sprd2Col1(1) = "Personal Tax"
      Case 2
        RevForm.vaSpread2.Text = "Machine Tools"
        Sprd2Col1(2) = "Machine Tools"
      Case 3
        RevForm.vaSpread2.Text = "Merchant Capital"
        Sprd2Col1(3) = "Merchant Capital"
      Case 4
        RevForm.vaSpread2.Text = "Farm Equipment"
        Sprd2Col1(4) = "Farm Equipment"
      Case 5
        RevForm.vaSpread2.Text = "Mobile Homes"
        Sprd2Col1(5) = "Mobile Homes"
      Case 6
        RevForm.vaSpread2.Text = "Interest"
        Sprd2Col1(6) = "Interest"
      Case 7
        RevForm.vaSpread2.Text = "Penalty"
        Sprd2Col1(7) = "Penalty"
      Case 8
        RevForm.vaSpread2.Text = QPTrim$(TaxMasterRec.POptRev1)
        TempPOptRev1 = TaxMasterRec.OptRev1
        Sprd2Col1(8) = QPTrim$(TaxMasterRec.POptRev1)
      Case 9
        RevForm.vaSpread2.Text = QPTrim$(TaxMasterRec.POptRev2)
        TempPOptRev2 = TaxMasterRec.OptRev2
        Sprd2Col1(9) = QPTrim$(TaxMasterRec.POptRev2)
      Case 10
        RevForm.vaSpread2.Text = QPTrim$(TaxMasterRec.POptRev3)
        TempPOptRev3 = TaxMasterRec.OptRev3
        Sprd2Col1(10) = QPTrim$(TaxMasterRec.POptRev3)
      Case Else
    End Select
    
    RevForm.vaSpread2.Col = 2
    Select Case x
    
      Case 1
        If TaxMasterRec.IntPersYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(1) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(1) = 0
        End If
        TempIntPersYN = TaxMasterRec.IntPersYN
      Case 2
        If TaxMasterRec.IntMTYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(2) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(2) = 0
        End If
        TempIntMTYN = TaxMasterRec.IntMTYN
      Case 3
        If TaxMasterRec.IntMCYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(3) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(3) = 0
        End If
        TempIntMCYN = TaxMasterRec.IntMCYN
      Case 4
        If TaxMasterRec.IntFEYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(4) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(4) = 0
        End If
        TempIntFEYN = TaxMasterRec.IntFEYN
      Case 5
        If TaxMasterRec.IntMHYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(5) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(5) = 0
        End If
        TempIntMHYN = TaxMasterRec.IntMHYN
      Case 6
        If TaxMasterRec.IntPIntYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(6) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(6) = 0
        End If
        TempIntPIntYN = TaxMasterRec.IntPIntYN
      Case 7
        If TaxMasterRec.IntPPenYN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(7) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(7) = 0
        End If
        TempIntPPenYN = TaxMasterRec.IntPPenYN
      Case 8
        If TaxMasterRec.IntPOpt1YN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(8) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(8) = 0
        End If
        TempPIntOpt1YN = TaxMasterRec.IntPOpt1YN
      Case 9
        If TaxMasterRec.IntPOpt2YN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(9) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(9) = 0
        End If
        TempIntPOpt2YN = TaxMasterRec.IntPOpt2YN
      Case 10
        If TaxMasterRec.IntPOpt3YN = "Y" Then
          RevForm.vaSpread2.Value = 1
          Sprd2Col2(10) = 1
        Else
          RevForm.vaSpread2.Value = 0
          Sprd2Col2(10) = 0
        End If
        TempIntPOpt3YN = TaxMasterRec.IntPOpt3YN
      Case Else
    End Select
  
    RevForm.vaSpread2.Col = 3
    Select Case x
      Case 1
        If TaxMasterRec.PersPayOrder < 1 Then
          RevForm.vaSpread2.Text = "1"
          Sprd2Col3(1) = 1
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.PersPayOrder)
          Sprd2Col3(1) = TaxMasterRec.PersPayOrder
        End If
      Case 2
        If TaxMasterRec.MTPayOrder < 1 Then
          RevForm.vaSpread2.Text = "2"
          Sprd2Col3(2) = 2
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.MTPayOrder)
          Sprd2Col3(2) = TaxMasterRec.MTPayOrder
        End If
      Case 3
        If TaxMasterRec.MCPayOrder < 1 Then
          RevForm.vaSpread2.Text = "3"
          Sprd2Col3(3) = 3
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.MCPayOrder)
          Sprd2Col3(3) = TaxMasterRec.MCPayOrder
        End If
      Case 4
        If TaxMasterRec.FEPayOrder < 1 Then
          RevForm.vaSpread2.Text = "4"
          Sprd2Col3(4) = 4
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.FEPayOrder)
          Sprd2Col3(4) = TaxMasterRec.FEPayOrder
        End If
      Case 5
        If TaxMasterRec.MHPayOrder < 1 Then
          RevForm.vaSpread2.Text = "5"
          Sprd2Col3(5) = 5
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.MHPayOrder)
          Sprd2Col3(5) = TaxMasterRec.MHPayOrder
        End If
      Case 6
        If TaxMasterRec.PIntPayOrder < 1 Then
          RevForm.vaSpread2.Text = "6"
          Sprd2Col3(6) = 6
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.PIntPayOrder)
          Sprd2Col3(6) = TaxMasterRec.PIntPayOrder
        End If
      Case 7
        If TaxMasterRec.PPenPayOrder < 1 Then
          RevForm.vaSpread2.Text = "7"
          Sprd2Col3(7) = 7
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.PPenPayOrder)
          Sprd2Col3(7) = TaxMasterRec.PPenPayOrder
        End If
      Case 8
        If TaxMasterRec.POpt1PayOrder < 1 Then
          RevForm.vaSpread2.Text = "8"
          Sprd2Col3(8) = 8
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.POpt1PayOrder)
          Sprd2Col3(8) = TaxMasterRec.POpt1PayOrder
        End If
      Case 9
        If TaxMasterRec.POpt2PayOrder < 1 Then
          RevForm.vaSpread2.Text = "9"
          Sprd2Col3(9) = 9
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.POpt2PayOrder)
          Sprd2Col3(9) = TaxMasterRec.POpt2PayOrder
        End If
      Case 10
        If TaxMasterRec.POpt3PayOrder < 1 Then
          RevForm.vaSpread2.Text = "10"
          Sprd2Col3(10) = 10
        Else
          RevForm.vaSpread2.Text = CStr(TaxMasterRec.POpt3PayOrder)
          Sprd2Col3(10) = TaxMasterRec.POpt3PayOrder
        End If
    End Select
  Next x
  
  fpcmbPPTRAYN.AddItem "Yes"
  fpcmbPPTRAYN.AddItem "No"
  If TaxMasterRec.PPTRAYN = "Y" Then
    fpcmbPPTRAYN.Text = "Yes"
  Else
    fpcmbPPTRAYN.Text = "No"
  End If
  TempPPTRAYN = TaxMasterRec.PPTRAYN
  
  fpcmbCyclesYN.AddItem "Yes"
  fpcmbCyclesYN.AddItem "No"
  If TaxMasterRec.UseCyclesYN = "Y" Then
    fpcmbCyclesYN.Text = "Yes"
  Else
    fpcmbCyclesYN.Text = "No"
  End If
  TempUseCyclesYN = TaxMasterRec.UseCyclesYN
  
  fpcmbCountyYN.AddItem "Yes"
  fpcmbCountyYN.AddItem "No"
  If TaxMasterRec.UseCountyYN = "Y" Then
    fpcmbCountyYN.Text = "Yes"
  Else
    fpcmbCountyYN.Text = "No"
  End If
  TempUseCountyYN = TaxMasterRec.UseCountyYN
  
  fpcmbRPSplitYN.AddItem "Yes"
  fpcmbRPSplitYN.AddItem "No"
  If TaxMasterRec.RealPersSplit <> "Y" Then
    fpcmbRPSplitYN.Text = "No"
    If fpcmbStateOfTax.Text = "NC" Then
      fpcmbRPSplitYN.Enabled = False
    Else
      fpcmbRPSplitYN.Enabled = True
    End If
  Else
    If fpcmbStateOfTax.Text = "VA" Or fpcmbStateOfTax.Text = "MD" Then
      fpcmbRPSplitYN.Text = "Yes"
      fpcmbRPSplitYN.Enabled = False
    Else
      fpcmbRPSplitYN.Enabled = True
      fpcmbRPSplitYN.Text = "Yes"
    End If
  End If
  
  TempRealPersSplit = TaxMasterRec.RealPersSplit
  If TaxMasterRec.LawChngDate > 0 Then
    fptxtLawChngDate.Text = MakeRegDate(TaxMasterRec.LawChngDate)
  Else
    fptxtLawChngDate.Text = "N/A"
  End If
  Call FixSpread
  
'  If Exist("C:\CPWork\ratetbls.dat") Then
'    Call TaxMsg(800, "Go to Tab 2 and click on the 'Revenue Setup' button to add optional revenues.")
''    RevForm.Show vbModal
'  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "LoadMe", Erl)
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
  
End Sub

Private Sub fpcmbAcctMeth_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbAcctMeth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAcctMeth.ListIndex = -1
  End If
  If fpcmbAcctMeth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLateFormat.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbCentDepYN_Change()
  If fpcmbCentDepYN.Text = "Yes" Then
    fptxtCentCash.Enabled = True
    fptxtCentSub.Enabled = True
    Label87.ForeColor = &H8000000E
    Label90.ForeColor = &H8000000E
    fptxtCentCash.BackColor = &H8000000E
    fptxtCentSub.BackColor = &H8000000E
  Else
    fptxtCentCash.Enabled = False
    fptxtCentSub.Enabled = False
    Label87.ForeColor = &H8000000B
    Label90.ForeColor = &H8000000B
    fptxtCentCash.BackColor = &H8000000B
    fptxtCentSub.BackColor = &H8000000B
  End If
End Sub

Private Sub fpcmbCentDepYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCentDepYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCentDepYN.ListIndex = -1
  End If
  If fpcmbCentDepYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtCentCash.Enabled = True Then
        fptxtCentCash.SetFocus
      Else
        fpcmbTaxBillFormat.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtDiscPPct.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbCountyYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCountyYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCountyYN.ListIndex = -1
  End If
  If fpcmbCountyYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbRPSplitYN.Enabled = True Then
        fpcmbRPSplitYN.SetFocus
      Else
        fptxtDiscRPct.SetFocus
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

Private Sub fpcmbCyclesYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCyclesYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCyclesYN.ListIndex = -1
  End If
  If fpcmbCyclesYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbCountyYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLateFormat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLateFormat.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLateFormat.ListIndex = -1
  End If
  If fpcmbLateFormat.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbNoInterYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbMinOptions_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMinOptions.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMinOptions.ListIndex = -1
  End If
  If fpcmbMinOptions.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtNameOfTaxAuth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbMultiYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMultiYear.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMultiYear.ListIndex = -1
  End If
  If fpcmbMultiYear.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtLawChngDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

'Private Sub fpcmbMultRevYN_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbMultRevYN.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbMultRevYN.ListIndex = -1
'  End If
'  If fpcmbMultRevYN.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fpcmbNoInterYN.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub

Private Sub fpcmbNoInterYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbNoInterYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbNoInterYN.ListIndex = -1
  End If
  If fpcmbNoInterYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtOverPayGL.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub


Private Sub fpcmbPPTRAYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPPTRAYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPPTRAYN.ListIndex = -1
  End If
  If fpcmbPPTRAYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpDSPPTRADisc.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcmbRPSplitYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRPSplitYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRPSplitYN.ListIndex = -1
  End If
  If fpcmbRPSplitYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtDiscRPct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStateOfTax_Change()
  If fpcmbStateOfTax.Text = "NC" Then
    fpcmbRPSplitYN.Action = ActionClear
    fpcmbRPSplitYN.Enabled = False
    fpcmbRPSplitYN.Text = "NA"
  ElseIf fpcmbStateOfTax.Text = "VA" Or fpcmbStateOfTax.Text = "MD" Then
    fpcmbRPSplitYN.Action = ActionClear
    fpcmbRPSplitYN.Enabled = False
    fpcmbRPSplitYN.Text = "Yes"
  Else
    fpcmbRPSplitYN.Action = ActionClear
    fpcmbRPSplitYN.Enabled = True
    fpcmbRPSplitYN.AddItem "Yes"
    fpcmbRPSplitYN.AddItem "No"
  End If

End Sub

Private Sub fpcmbStateOfTax_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbStateOfTax.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStateOfTax.ListIndex = -1
  End If
  If fpcmbStateOfTax.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtCurrYrRIntRate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxBillFormat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxBillFormat.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxBillFormat.ListIndex = -1
  End If
  If fpcmbLateFormat.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbAcctMeth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Function Check4Changes() As Boolean
  Dim TaxRec As TaxMasterType
  Dim TMHandle As Integer
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThisDbl As Double
  Dim OptStr As String
  Dim OptInt As Integer
  Dim TabNum As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  x = 0
  Check4Changes = False
  If Exist("TAXSETUP.DAT") Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxRec
  Else
    Exit Function
  End If
  
  '----Tab 0----------------------------------------------
  Set ThisControl = fptxtNameOfTaxAuth
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.Name)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Name of Taxing Authority' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.Name = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "The Name of Taxing Authority has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtAdd1
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.Add1)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Address 1' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.Add1 = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Address #1 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtAdd2
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.Add2)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Address 2' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.Add2 = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Address #2 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCity
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.City)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'City' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.City = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "The City has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtState
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.TownState)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'State' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.TownState = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "The State has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtZip
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.Zip)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Zip Code' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.Zip = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Zip Code has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbStateOfTax
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.TaxSt)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'State of Tax' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.TaxSt = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "State of Tax has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCustOptSrch
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.OptSrchCust)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'For Customer' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchCust = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "For Customer has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPropOptSrch
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.OptSrchProp)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'For Property' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchCust = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "For Property has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPersOptSrch
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.OptSrchPers)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'For Personal' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchPers = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "For Personal has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrYrRIntRate
  TabNum = 0
  ThisDbl = TaxRec.CurrRYrIntInUse
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Current Year Real Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.CurrRYrIntInUse = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Current Year Real Interest Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrYrPIntRate
  TabNum = 0
  ThisDbl = TaxRec.CurrPYrIntInUse
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Current Year Personal Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.CurrRYrIntInUse = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Current Year Personal Interest Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrRYear
  TabNum = 0
  OptInt = TaxRec.RTaxYear
  If CInt(ThisControl.Text) <> OptInt Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Real Current Tax Year' field has been changed from " + CStr(OptInt) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.RTaxYear = CInt(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Real Current Tax Year has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrPYear
  TabNum = 0
  OptInt = TaxRec.PTaxYear
  If CInt(ThisControl.Text) <> OptInt Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Personal Current Tax Year' field has been changed from " + CStr(OptInt) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.PTaxYear = CInt(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Personal Current Tax Year has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPastYearIntRate
  TabNum = 0
  ThisDbl = TaxRec.PastYrInt
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Past Year Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.PastYrInt = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Past Year Interest Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtMinTaxAmt
  TabNum = 0
  ThisDbl = TaxRec.MinBill
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Minimum Tax Amount' field has been changed from " + Using("$##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.MinBill = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Minimum Tax Amount has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbMinOptions
  TabNum = 0
  ThisDesc = CStr(TaxRec.MinTxOpt)
  If InStr(ThisControl.Text, ThisDesc) = 0 Then
    Select Case TaxRec.MinTxOpt
      Case 0
        ThisDesc = "(0)"
      Case 1
        ThisDesc = "(1)"
      Case 2
        ThisDesc = "(2)"
      Case Else
        ThisDesc = "(0)"
    End Select
    If InStr(ThisControl.Text, "(0)") > 0 Then
      OptStr = "(0)"
    ElseIf InStr(ThisControl.Text, "(1)") > 0 Then
      OptStr = "(1)"
    ElseIf InStr(ThisControl.Text, "(2)") > 0 Then
      OptStr = "(2)"
    Else
      OptStr = "(0)"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Minimum Tax Options' field has been changed from " + ThisDesc + " to " + OptStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If OptStr = "(0)" Then
        TaxRec.MinTxOpt = 0
      ElseIf OptStr = "(1)" Then
        TaxRec.MinTxOpt = 1
      ElseIf OptStr = "(2)" Then
        TaxRec.MinTxOpt = 2
      Else
        TaxRec.MinTxOpt = 0
      End If
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "The minimum tax option has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  '----Tab 1----------------------------------------------
  
'  Set ThisControl = fptxtPenaltyRate
'  TabNum = 1
'  ThisDbl = TaxRec.PenPct
'  If CDbl(ThisControl.Text) <> ThisDbl Then
'    frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalty Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmVATaxMsgW4Opts.Label1.Top = 575
'    frmVATaxMsgW4Opts.Show vbModal
'    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW4Opts
'    If choice = "save" Then
'      TaxRec.PenPct = CDbl(ThisControl.Text)
'      Put TMHandle, 1, TaxRec
'      Call Savemsg(900, "Penalty Interest Rate has been saved successfully.")
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
   Set ThisControl = fpcmbPPTRAYN
  TabNum = 1
  ThisDesc = TaxRec.PPTRAYN
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'PPTRA Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.PPTRAYN = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "PPTRA Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fpDSPPTRADisc
  TabNum = 1
  ThisDbl = TaxRec.PPTRADisc
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'PPTRA Discount Pct' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.PPTRADisc = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "PPTRA Discount Pct has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
 Set ThisControl = fpcmbCentDepYN
  TabNum = 1
  ThisDesc = TaxRec.CntrlDepYN
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Central Depository Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.CntrlDepYN = ThisControl.Text
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Central Depository Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtCentCash
  TabNum = 1
  ThisDesc = QPTrim$(TaxRec.CDCashGL)
  If QPTrim(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Central Depository Cash G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtCentCash.Text)) = False Then
        frmVATaxMsgWOpts.Label1.Caption = "The Central Depository Cash G/L number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmVATaxMsgWOpts.Label1.Top = 600
        frmVATaxMsgWOpts.Show vbModal
        If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmVATaxMsgWOpts
          MainLog ("Warning: User issued warning that the central depository cash GL number " + QPTrim$(fptxtCentCash.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmVATaxMsgWOpts
          Close
          vaTabPro1.ActiveTab = 1
          If fptxtCentCash.Enabled = True Then
            fptxtCentCash.SetFocus
          Else
            fpcmbCentDepYN.SetFocus
          End If
          Check4Changes = True
          Exit Function
        End If
      End If
      TaxRec.CDCashGL = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Central Depository Cash G/L number has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fptxtCentSub
  TabNum = 1
  ThisDesc = QPTrim$(TaxRec.CDSubGL)
  If QPTrim(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Central Depository Sub G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtCentSub.Text)) = False Then
        frmVATaxMsgWOpts.Label1.Caption = "The Central Depository Sub G/L number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmVATaxMsgWOpts.Label1.Top = 600
        frmVATaxMsgWOpts.Show vbModal
        If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmVATaxMsgWOpts
          MainLog ("Warning: User issued warning that the central depository sub GL number " + QPTrim$(fptxtCentSub.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmVATaxMsgWOpts
          Close
          vaTabPro1.ActiveTab = 1
          If fptxtCentSub.Enabled = True Then
            fptxtCentSub.SetFocus
          Else
            fpcmbCentDepYN.SetFocus
          End If
          Check4Changes = True
          Exit Function
        End If
      End If
      TaxRec.CDSubGL = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Central Depository Sub G/L number has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fpcmbNoInterYN
  TabNum = 1
  ThisDesc = TaxRec.WarnInt
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'No Interest Warning Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.WarnInt = ThisControl.Text
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "No Interest Warning Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fpcmbMultiYear
  TabNum = 1
  ThisDesc = TaxRec.MultiYear
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Multi Year' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.MultiYear = ThisControl.Text
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Multi Year has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  Set ThisControl = fpcmbAcctMeth
  TabNum = 1
  ThisDesc = QPTrim$(TaxRec.AcctgMethod)
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    Select Case QPTrim$(TaxRec.AcctgMethod)
      Case "N"
        ThisDesc = "NONE"
      Case "C"
        ThisDesc = "CASH"
      Case "M"
        ThisDesc = "MODIFIED ACCRUAL"
      Case "A"
        ThisDesc = "ACCRUAL"
      Case Else
        ThisDesc = "NONE"
    End Select
    Select Case Mid(ThisControl.Text, 1, 1)
      Case "N"
        OptStr = "NONE"
      Case "C"
        OptStr = "CASH"
      Case "M"
        OptStr = "MODIFIED ACCRUAL"
      Case "A"
        OptStr = "ACCRUAL"
      Case Else
        OptStr = "NONE"
    End Select
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Accounting Method' field has been changed from " + ThisDesc + " to " + OptStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.AcctgMethod = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Accounting Method has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
 
  Select Case TaxRec.TaxForm
    Case 30000
      ThisDesc = "STANDARD"
'    Case 21837
'      ThisDesc = "MULTI-PART"
'    Case 20304
'      ThisDesc = "POSTCARD"
    Case 16716
      ThisDesc = "LASER"
    Case 20000
      ThisDesc = "EXPORT REAL"
    Case 20001
      ThisDesc = "EXPORT PERSONAL"
    Case 20002
      ThisDesc = "LASER ITEMIZED"
    Case 20003
      ThisDesc = "MDLTWN"
    Case 20004
      ThisDesc = "CDRBLUFF"
    Case Else
      ThisDesc = "UNKNOWN"
  End Select
  
  Select Case QPTrim$(fpcmbTaxBillFormat.Text)
    Case "STANDARD"
      OptInt = 30000
'    Case "MULTI-PART"
'      OptInt = 21837
'    Case "POSTCARD"
'      OptInt = 20304
    Case "LASER"
      OptInt = 16716
    Case "EXPORT REAL"
      OptInt = 20000
    Case "EXPORT PERSONAL"
      OptInt = 20001
    Case 20002
      OptInt = "LASER ITEMIZED"
    Case 20003
      OptInt = "MDLTWN"
    Case 20004
      OptInt = "CDRBLUFF"
    Case Else
      OptInt = 0
  End Select
  
  Set ThisControl = fpcmbTaxBillFormat
  TabNum = 1
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Tax Bill Format' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.TaxForm = OptInt
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Tax Bill Format has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  '***************Late Bill Format goes here********************
  Set ThisControl = fpcmbLateFormat
  TabNum = 1
  OptInt = TaxRec.LateForm
  If CInt(Mid(ThisControl.Text, 1, 1)) <> OptInt Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Late Bill Format' field has been changed from " + CStr(OptInt) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.LateForm = CInt(Mid(ThisControl.Text, 1, 1))
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Late Bill Format has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtOverPayGL
  TabNum = 1
  ThisDesc = QPTrim$(TaxRec.OverPayGLNum)
  If QPTrim(ThisControl.Text) <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Overpayment G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtOverPayGL.Text)) = False Then
        frmVATaxMsgWOpts.Label1.Caption = "The Overpayment GL number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmVATaxMsgWOpts.Label1.Top = 600
        frmVATaxMsgWOpts.Show vbModal
        If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmVATaxMsgWOpts
          MainLog ("Warning: User issued warning that the overpayment GL number " + QPTrim$(fptxtOverPayGL.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmVATaxMsgWOpts
          Close
          vaTabPro1.ActiveTab = 1
          fptxtOverPayGL.SetFocus
          Check4Changes = True
          Exit Function
        End If
      End If
      TaxRec.OverPayGLNum = ThisControl.Text
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Overpayment G/L number has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
'  Set ThisControl = fpcmbMultRevYN
'  TabNum = 1
'  ThisDesc = TaxRec.PriorYrMltRevYN
'  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
'    If ThisDesc = "N" Then
'      ThisDesc = "No"
'    ElseIf ThisDesc = "Y" Then
'      ThisDesc = "Yes"
'    End If
'    frmVATaxMsgW4Opts.Label1.Caption = "The 'Do you use multiple revenue accounts for prior years Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmVATaxMsgW4Opts.Label1.Top = 575
'    frmVATaxMsgW4Opts.Show vbModal
'    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW4Opts
'    If choice = "save" Then
'      TaxRec.PriorYrMltRevYN = ThisControl.Text
'      Put TMHandle, 1, TaxRec
'      Call Savemsg(900, "Do you use multiple revenue accounts for prior years Y/N? has been saved successfully.")
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fptxtDiscRPct
  TabNum = 1
  ThisDbl = TaxRec.DisRPct
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Real Discount Percentage' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.DisRPct = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Real Discount Percentage has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDiscPPct
  TabNum = 1
  ThisDbl = TaxRec.DisPPct
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Personal Discount Percentage' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.DisPPct = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Personal Discount Percentage has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Dim SpreadText As String
  Dim ThisCol As Integer
  
  ThisCol = 1
  RevForm.vaSpread1.Col = ThisCol
  For x = 6 To 8
    RevForm.vaSpread1.Row = x
    If x = 6 Then
      SpreadText = QPTrim$(RevForm.vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev1)
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #1' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmVATaxMsgW4Opts.Label1.Top = 375
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.OptRev1 = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Optional revenue #1's name was saved successfully.")
          MainLog ("User warned that changing the name of Optional Revenue #1 from " + ThisDesc + " to " + SpreadText + " has reporting consequences. The user elected to save the change anyway.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 7 Then
      SpreadText = QPTrim$(RevForm.vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev2)
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #2' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmVATaxMsgW4Opts.Label1.Top = 375
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.OptRev2 = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Optional Revenue #2's name was saved successfully.")
          MainLog ("User warned that changing the name of Optional Revenue #2 from " + ThisDesc + " to " + SpreadText + " has reporting consequences. The user elected to save the change anyway.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 8 Then
      SpreadText = QPTrim$(RevForm.vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev3)
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #3' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmVATaxMsgW4Opts.Label1.Top = 375
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.OptRev3 = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Optional Revenue #3's name was saved successfully.")
          MainLog ("User warned that changing the name of Optional Revenue #3 from " + ThisDesc + " to " + SpreadText + " has reporting consequences. The user elected to save the change anyway.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
  Next x
  
  Dim SpreadText2 As String
  
  ThisCol = 2
  RevForm.vaSpread1.Col = ThisCol
  For x = 1 To 8
    RevForm.vaSpread1.Row = x
    If RevForm.vaSpread1.Text = "1" Then
      SpreadText = "Y"
    Else
      SpreadText = "N"
    End If
    RevForm.vaSpread1.Col = 1
    SpreadText2 = QPTrim$(RevForm.vaSpread1.Text)
    RevForm.vaSpread1.Col = ThisCol
    If x = 1 Then
      ThisDesc = TaxRec.IntIntYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntIntYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 2 Then
      ThisDesc = TaxRec.IntAdvYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntAdvYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 3 Then
      ThisDesc = TaxRec.IntLateLstYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntLateLstYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 4 Then
      ThisDesc = TaxRec.IntPenaltyYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntPenaltyYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 5 Then
      ThisDesc = TaxRec.IntPrncTaxYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntPrncTaxYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 6 Then
      ThisDesc = TaxRec.IntOpt1YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntOpt1YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 7 Then
      ThisDesc = TaxRec.IntOpt2YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntOpt2YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 8 Then
      ThisDesc = TaxRec.IntOpt3YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntOpt3YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
  Next x
  
  ThisCol = 3
  RevForm.vaSpread1.Col = ThisCol
  For x = 6 To 8
    RevForm.vaSpread1.Row = x
    If RevForm.vaSpread1.Text = "1" Then
      SpreadText = "Y"
      If TaxRec.PenIdx <> x Then
        ThisDesc = "N"
        RevForm.vaSpread1.Col = 1
        SpreadText2 = QPTrim$(RevForm.vaSpread1.Text)
        RevForm.vaSpread1.Col = ThisCol
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalty Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          MainLog ("The penalty revenue selection has changed from revenue # " + CStr(TaxRec.PenIdx) + " to revenue # " + CStr(x) + " and saved.")
          TaxRec.PenIdx = x
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalty Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
  Next x
         
  ThisCol = 4
  RevForm.vaSpread1.Col = ThisCol
  For x = 1 To 8
    RevForm.vaSpread1.Row = x
    If RevForm.vaSpread1.Text = "1" Then
      SpreadText = "Y"
    Else
      SpreadText = "N"
    End If
    RevForm.vaSpread1.Col = 1
    SpreadText2 = QPTrim$(RevForm.vaSpread1.Text)
    RevForm.vaSpread1.Col = ThisCol
    If x = 1 Then
      ThisDesc = TaxRec.PenIntYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenIntYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 2 Then
      ThisDesc = TaxRec.PenAdvYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenAdvYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 3 Then
      ThisDesc = TaxRec.PenLateLstYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenLateLstYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 4 Then
      ThisDesc = TaxRec.PenPenaltyYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenPenaltyYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 5 Then
      ThisDesc = TaxRec.PenPrncTaxYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenPrncTaxYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 6 Then
      ThisDesc = TaxRec.PenOpt1YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenOpt1YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 7 Then
      ThisDesc = TaxRec.PenOpt2YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenOpt2YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 8 Then
      ThisDesc = TaxRec.PenOpt3YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmVATaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmVATaxMsgW4Opts.Label1.Top = 575
        frmVATaxMsgW4Opts.Show vbModal
        choice = frmVATaxMsgW4Opts.fptxtChoice.Text
        Unload frmVATaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenOpt3YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
  Next x
  
  x = 0
  
  Set ThisControl = fpcmbCyclesYN
  TabNum = 1
  ThisDesc = TaxRec.UseCyclesYN
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Use Billing Cycles Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.UseCyclesYN = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Use Billing Cycles Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
    
  Set ThisControl = fpcmbCountyYN
  TabNum = 1
  ThisDesc = TaxRec.UseCountyYN
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Use County Billing Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.UseCountyYN = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Use County Billing Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  If fpcmbRPSplitYN.Enabled = True Then
    Set ThisControl = fpcmbRPSplitYN
    TabNum = 1
    ThisDesc = TaxRec.RealPersSplit
    If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
      If ThisDesc = "N" Then
        ThisDesc = "No"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Yes"
      End If
      frmVATaxMsgW4Opts.Label1.Caption = "The 'Use Real/Personal Split Billing Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmVATaxMsgW4Opts.Label1.Top = 575
      frmVATaxMsgW4Opts.Show vbModal
      choice = frmVATaxMsgW4Opts.fptxtChoice.Text
      Unload frmVATaxMsgW4Opts
      If choice = "save" Then
        TaxRec.RealPersSplit = Mid(ThisControl.Text, 1, 1)
        Put TMHandle, 1, TaxRec
        Call Savemsg(900, "Use Real/Personal Split Billing Y/N? has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Set ThisControl = fpCurrMaxVehAmt
  TabNum = 1
  ThisDbl = TaxRec.MaxVehTaxVal
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Maximum Vehicle Tax Value' field has been changed from " + QPTrim$(Using("$###,##0.00", ThisDbl)) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.MaxVehTaxVal = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Maximum Vehicle Tax Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMinVehAmt
  TabNum = 1
  ThisDbl = TaxRec.MinVehTaxVal
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Minimum Vehicle Tax Value' field has been changed from " + QPTrim$(Using("$###,##0.00", ThisDbl)) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    frmVATaxMsgW4Opts.Show vbModal
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      TaxRec.MinVehTaxVal = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Minimum Vehicle Tax Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Close TMHandle
  
  Exit Function
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close TMHandle
        frmVATaxBillSetUpMenu.Show
        DoEvents
        Unload RevForm
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        vaTabPro1.ActiveTab = TabNum
        If x > 0 Then
          RevForm.Show vbModal
          RevForm.vaSpread1.SetActiveCell ThisCol, x
        Else
          ThisControl.SetFocus
        End If
        Close TMHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "Check4Changes", Erl)
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

Private Sub LogSaves()
  Dim TaxRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TempStr$, TempSave$
  Dim TempDbl As Double
  Dim TempSaveDbl As Double
  Dim TempInt As Integer
  Dim TempSaveInt As Integer
  Dim ThisZip$
  Dim ThatZip$
  
  On Local Error Resume Next
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxRec
  Close TMHandle
  
  TempStr = QPTrim$(TempName)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Name)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Name of Taxing Authority was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempADD1)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Add1)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Address #1 was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempADD2)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Add2)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Address #2 was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempCity)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.City)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: City was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempTownState)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.TownState)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The town's state was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempZip)
  ThisZip = ReplaceString(TempStr, "-", "")
  If QPTrim$(ThisZip) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Zip)
  ThatZip = ReplaceString(TempSave, "-", "")
  If QPTrim$(ThatZip) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Zip Code was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempTaxSt)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.TaxSt)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The tax state was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchCust)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchCust)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The optional customer search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchProp)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchProp)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The optional property search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchPers)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchPers)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The optional personal search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempDbl = TempCurrYrRInt
  TempSaveDbl = TaxRec.CurrRYrIntInUse
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The current year's real tax percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempDbl = TempCurrYrPInt
  TempSaveDbl = TaxRec.CurrPYrIntInUse
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The current year's personal tax percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempInt = TempRTaxYear
  TempSaveInt = TaxRec.RTaxYear
  If TempInt <> TempSaveInt Then
    MainLog ("frmVATaxSystemSetup: The real current tax year was changed from " + CStr(TempInt) + " to " + CStr(TempSaveInt) + " and saved.")
  End If
    
  TempInt = TempPTaxYear
  TempSaveInt = TaxRec.PTaxYear
  If TempInt <> TempSaveInt Then
    MainLog ("frmVATaxSystemSetup: The personal current tax year was changed from " + CStr(TempInt) + " to " + CStr(TempSaveInt) + " and saved.")
  End If
    
  TempDbl = TempPastYrInt
  TempSaveDbl = TaxRec.PastYrInt
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The past year's tax percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempDbl = TempPenPct
  TempSaveDbl = TaxRec.PenPct
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The penalty percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  Select Case TempTaxForm
    Case 30000
      TempStr = "STANDARD"
'    Case 20304
'      TempStr = "POSTCARD"
'    Case 21837
'      TempStr = "MULTI-PART"
    Case 16716
      TempStr = "LASER"
    Case Else
      TempStr = "UNKNOWN"
  End Select
  Select Case TaxRec.TaxForm
    Case 30000
      TempSave = "STANDARD"
'    Case 20304
'      TempSave = "POSTCARD"
'    Case 21837
'      TempSave = "MULTI-PART"
    Case 16716
      TempSave = "LASER"
    Case Else
      TempSave = "UNKNOWN"
  End Select
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The tax bill format was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  Select Case TempMinTxOpt
    Case 0
      TempStr = "(0) No special..."
    Case 1
      TempStr = "(1) Charge no tax..."
    Case 2
      TempStr = "(2) Charge minimum..."
    Case Else
      TempStr = "UNKNOWN"
  End Select
  Select Case TaxRec.MinTxOpt
    Case 0
      TempSave = "(0) No special..."
    Case 1
      TempSave = "(1) Charge no tax..."
    Case 2
      TempSave = "(2) Charge minimum..."
    Case Else
      TempSave = "UNKNOWN"
  End Select
  
  If TempMinTxOpt = "" Then TempMinTxOpt = "0"
  TempInt = TempMinTxOpt
  TempSaveInt = TaxRec.MinTxOpt
  If TempInt <> TempSaveInt Then
    MainLog ("frmVATaxSystemSetup: The minimum tax option was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempDbl = TempMinTxPct
  TempSaveDbl = TaxRec.MinBill
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The minimum tax amount was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  Select Case TempAcctgMethod
    Case "N"
      TempStr = "NONE"
    Case "A"
      TempStr = "ACCRUAL"
    Case "C"
      TempStr = "CASH"
    Case "M"
      TempStr = "MODIFIED ACCRUAL"
    Case Else
      TempStr = "UNKNOWN"
  End Select
  
  Select Case TaxRec.AcctgMethod
    Case "N"
      TempSave = "NONE"
    Case "A"
      TempSave = "ACCRUAL"
    Case "C"
      TempSave = "CASH"
    Case "M"
      TempSave = "MODIFIED ACCRUAL"
    Case Else
      TempSave = "UNKNOWN"
  End Select
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: The accounting method was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  If TempDisRPct = "" Then TempDisRPct = "0"
  TempDbl = TempDisRPct
  TempSaveDbl = TaxRec.DisRPct
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The real discount amount was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  If TempDisPPct = "" Then TempDisPPct = "0"
  TempDbl = TempDisPPct
  TempSaveDbl = TaxRec.DisPPct
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The personal discount amount was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  TempStr = TempCntrlDepYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.CntrlDepYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Central Depository Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempCDCashGL)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.CDCashGL)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Central Depository Cash G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempCDSubGL)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.CDSubGL)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Central Depository Sub G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
'  TempStr = TempPriorYrMltRevYN
'  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
'  TempSave = TaxRec.PriorYrMltRevYN
'  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
'  If TempStr <> TempSave Then
'    MainLog ("frmVATaxSystemSetup: Do you use multiple revenue accounts for prior years Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
'  End If
  
  TempStr = QPTrim$(TempOverPayGLNum)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OverPayGLNum)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Overpayment G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOverPayGLNum)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OverPayGLNum)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Overpayment G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  '-------------------------------------------------------------------------
  
  TempStr = TempIntPrncTaxYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntPrncTaxYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest on This Rev for Principle was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempIntIntYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntIntYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest on Rev for Interest Accrued was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempIntAdvYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntAdvYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest on Rev for Advertising was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempIntLateLstYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntLateLstYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest on Rev for Late Listing was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempIntPenaltyYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntPenaltyYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest on Rev for Penalty was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  '-------------------------------------------------------------------------
  TempStr = TempPenPrncTaxYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenPrncTaxYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for Principle was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenIntYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenIntYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for Interest Accrued was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenAdvYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenAdvYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for Advertising was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenLateLstYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenLateLstYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for Late Listing was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempPenPenaltyYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenPenaltyYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for Penalty was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempWarnInt
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.WarnInt
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: No Interest Warning Y/N was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev1)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev1)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Revenue Description for Optional Revenue #1 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  TempStr = TempIntOpt1YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntOpt1YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest for This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
'  RevForm.vaSpread1.Col = 1
'  RevForm.vaSpread1.Row = 5
  TempStr = TempPenOpt1YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt1YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev2)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev2)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Revenue Description for Optional Revenue #2 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  TempStr = TempIntOpt2YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntOpt2YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest for This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
'  RevForm.vaSpread1.Col = 1
'  RevForm.vaSpread1.Row = 6
  TempStr = TempPenOpt2YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt2YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev3)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev3)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Revenue Description for Optional Revenue #3 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  TempStr = TempIntOpt3YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.IntOpt3YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Interest for This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempPenOpt3YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt3YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Penalize This Rev for " + QPTrim$(RevForm.vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  If TaxRec.PenIdx > 0 Then
    RevForm.vaSpread1.Col = 1
    RevForm.vaSpread1.Row = TaxRec.PenIdx
    MainLog ("frmVATaxSystemSetup: Penalty revenue saved as " + QPTrim$(RevForm.vaSpread1.Text) + " on row # " + CStr(TaxRec.PenIdx) + ".")
  Else
    MainLog ("frmVATaxSystemSetup: No penalty revenue index was saved.")
  End If
  
  TempStr = TempUseCyclesYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.UseCyclesYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Use Billing Cycles Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempUseCountyYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.UseCountyYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmVATaxSystemSetup: Use County Billing Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempInt = TempRealPersSplit
  TempSaveInt = TaxRec.RealPersSplit
  If TempInt <> TempSaveInt Then
    MainLog ("frmVATaxSystemSetup: Use Real/Pers Split Billing Y/N? was changed from " + CStr(TempInt) + " to " + CStr(TempSaveInt) + " and saved.")
  End If
  
  TempDbl = TempMaxVehVal
  TempSaveDbl = TaxRec.MaxVehTaxVal
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The maximum vehicle tax value was changed from " + Using("$###,##0.00", TempDbl) + " to " + Using("$###,##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempDbl = TempMinVehVal
  TempSaveDbl = TaxRec.MinVehTaxVal
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmVATaxSystemSetup: The minimum vehicle tax value was changed from " + Using("$###,##0.00", TempDbl) + " to " + Using("$###,##0.00", TempSaveDbl) + " and saved.")
  End If
  
  TempInt = TempMultiYear
  TempSaveInt = TaxRec.MultiYear
  If TempInt <> TempSaveInt Then
    MainLog ("frmVATaxSystemSetup: The multi year tax value was changed from " + Using("##", TempInt) + " to " + Using("##", TempSaveInt) + " and saved.")
  End If
    
    
End Sub

Private Sub fpCurrMinVehAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbPPTRAYN.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpCurrMaxVehAmt.SetFocus
  End If
End Sub

Private Sub fpDSPPTRADisc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbCyclesYN.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpcmbPPTRAYN.SetFocus
  End If
End Sub

Private Sub fpListTownships_DblClick()
  cmdAddTownship.Text = "Edit Township"
  TSListIdx = fpListTownships.ListIndex
  fptxtTownShipName.Text = QPTrim$(fpListTownships.Text)
End Sub

Private Sub fptxtCentCash_DblClick(Button As Integer)
  fptxtCentCash.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtCentSub_DblClick(Button As Integer)
  fptxtCentSub.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtCustOptSrch_Change()
  If InStr(fptxtCustOptSrch.Text, " ORDER") Then
    If TaxMsgWOpts(750, "Using the word 'Order' in the description will be displayed as " + fptxtCustOptSrch.Text + " Order in the Print Order drop down boxes. If this is OK then press F10 to approve. Otherwise press ESC to return and edit.", "F10 Approve", "ESC Edit") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCustOptSrch.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub fptxtDiscPct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbCentDepYN.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpcmbRPSplitYN.Enabled = True Then
      fpcmbRPSplitYN.SetFocus
    Else
      fpcmbCountyYN.SetFocus
    End If
  End If
End Sub

Private Sub fptxtDiscPPct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbCentDepYN.SetFocus
  ElseIf KeyCode = vbKeyDown Then
    fptxtDiscRPct.SetFocus
  End If
End Sub

Private Sub fptxtOverPayGL_DblClick(Button As Integer)
  fptxtOverPayGL.Text = Clipboard.GetText
  frmVATaxGLList.ZOrder 0
End Sub

Private Sub fptxtOverPayGL_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  Dim ThisGL$
  
  If QPTrim$(fptxtOverPayGL.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
  ThisGL$ = fptxtOverPayGL.Text
  fptxtOverPayGL.Text = ReplaceString(fptxtOverPayGL.Text, "-", "")
  ThisLen = Len(QPTrim$(fptxtOverPayGL.Text))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtOverPayGL.Text = ThisGL$
      fptxtOverPayGL.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fptxtOverPayGL.Text = ThisGL$
      Exit Sub
    End If
  End If
  fptxtOverPayGL.Text = AddDashesToGLNumber(fptxtOverPayGL.Text, Fund, Dept, Detail)
End Sub


'Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
'  Dim PenCnt As Integer
'  Dim x As Integer
'  Dim CntPens As Integer
'  Dim Thisx As Integer
'
'  On Error GoTo ERRORSTUFF
'
'  StrEmpty = False
'  RevForm.vaSpread1.Col = 1
'  RevForm.vaSpread1.Row = Row
'  If QPTrim$(RevForm.vaSpread1.Text) = "" Then
'    StrEmpty = True
'  End If
'  RevForm.vaSpread1.Col = 2
'  If RevForm.vaSpread1.Text = "1" And StrEmpty = True Then
'    Call TaxMsg(800, "This row contains an unused optional revenue. Setting interest is not allowed for this row.")
'    RevForm.vaSpread1.Text = "0"
'    vaTabPro1.ActiveTab = 1
'    RevForm.vaSpread1.SetFocus
'    RevForm.vaSpread1.SetActiveCell 2, Row
'    Exit Sub
'  End If
'
'  If Col <> 3 Then Exit Sub
'  CntPens = 0
'  Thisx = 0
'  RevForm.vaSpread1.Col = 3
'  For x = 5 To 7
'    RevForm.vaSpread1.Row = x
'    If RevForm.vaSpread1.Text = "1" Then
'      CntPens = CntPens + 1
'      Thisx = x
'    End If
'  Next x
'
'  If CntPens > 1 Then
'    Call TaxMsg(800, "ERROR: Only one optional revenue can be earmarked as the penalty revenue. Please review your penalty selections and select only one penalty revenue.")
'    vaTabPro1.ActiveTab = 1
'    RevForm.vaSpread1.SetActiveCell 3, Thisx
'    Exit Sub
'  End If
'
'  RevForm.vaSpread1.Col = Col
'  RevForm.vaSpread1.Row = Row
'
'
'  If RevForm.vaSpread1.Text = "1" And PenIdx <> Row Then
'    RevForm.vaSpread1.Col = 1
'    If QPTrim$(RevForm.vaSpread1.Text) = "" Then
'      Call TaxMsg(800, "The penalty revenue source has been assigned to row " + CStr(Row) + ". Please enter a penalty description.")
'      vaTabPro1.ActiveTab = 1
'      RevForm.vaSpread1.SetActiveCell 1, Row
'    End If
'    RevForm.vaSpread1.Col = Col
'    If PenIdx > 0 Then
'      RevForm.vaSpread1.Row = PenIdx
'      RevForm.vaSpread1.Value = "0"
'    End If
'    PenIdx = Row
'  End If
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "RevForm.vaSpread1_Change", Erl)
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
'    Terminate
'
'End Sub

Private Sub vaTabPro1_GotFocus()
  If StrEmpty = True Then 'StrEmpty = True when user checks interest
  'for an unused opt revenue
    StrEmpty = False
    Exit Sub
  End If
  If vaTabPro1.ActiveTab = 1 Then
    If fpcmbCentDepYN.Enabled = True Then
      fpcmbCentDepYN.SetFocus
    End If
  ElseIf vaTabPro1.ActiveTab = 0 Then
    If fptxtNameOfTaxAuth.Enabled = True Then
      fptxtNameOfTaxAuth.SetFocus
    End If
  End If
End Sub

Private Function VerifyGLNum(GLNum$) As Boolean
   Dim IdxRec As JGLAcctIdxType
   Dim GLIdxNum$
   Dim IdxHandle As Integer
   Dim IdxCnt As Integer
   Dim x As Integer, y As Integer
   Dim GLRec As GLAcctRecType
   Dim GLHandle As Integer
   Dim GLCnt As Integer
   Dim CheckThis$
   
   On Error GoTo ERRORSTUFF
   
   VerifyGLNum = True
   
   If Not Exist("GLACCT.IDX") Then
     Call TaxMsg(900, "Unable to locate 'GLACCT.IDX'. General Ledger numbers cannot be verified.")
     Exit Function
   End If
   
   OpenGLIdxFile IdxHandle, IdxCnt
   
   ReDim IdxRecs(1 To IdxCnt) As Integer
   If IdxCnt = 0 Then
     Close
     Exit Function
   End If
   
   For x = 1 To IdxCnt
     Get IdxHandle, x, IdxRec
     IdxRecs(x) = IdxRec.RecNo
   Next x
   Close IdxHandle
   
   If Not Exist("GLACCT.DAT") Then
     Call TaxMsg(900, "Unable to locate 'GLACCT.DAT'. General Ledger numbers cannot be verified.")
     Exit Function
   End If
   
   OpenGLAcctFile GLHandle, GLCnt
   If GLCnt = 0 Then
     Close GLHandle
     Exit Function
   End If
   
   CheckThis = QPTrim$(fptxtOverPayGL.Text)
   For x = 1 To IdxCnt
     If IdxRecs(x) <> 0 Then
       Get GLHandle, IdxRecs(x), GLRec
       If GLRec.Deleted Then GoTo SkipIt
       If CheckThis = QPTrim$(GLRec.Num) Then
         Exit For
       End If
     End If
SkipIt:
   Next x
   Close GLHandle
   
   If x > IdxCnt Then
     VerifyGLNum = False
   End If
   
   Exit Function
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxSystemSetup", "VerifyGLNum", Erl)
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
    frmVATaxBillSetUpMenu.Show
    DoEvents
    Unload RevForm
    Unload Me

End Function

Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim coladj As Integer
  Dim x As Integer, y As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 21
        coladj = 10
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 12
          Next y
        Next x
        RevForm.vaSpread1.RowHeight(-1) = 27.5
        RevForm.vaSpread1.RowHeight(0) = 27.5
      Else
        COne = -2 '11.25
        coladj = 3 '3.45
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread1.Col = 2
        RevForm.vaSpread1.Row = 0
        RevForm.vaSpread1.FontName = "Tahoma"
        RevForm.vaSpread1.FontSize = 10
        RevForm.vaSpread1.Text = RevForm.vaSpread1.Text
        RevForm.vaSpread1.RowHeight(-1) = 18 '23.5
        RevForm.vaSpread1.RowHeight(0) = 18 '23.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 15
        coladj = 7
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 14
          Next y
        Next x
        RevForm.vaSpread1.RowHeight(0) = 24
        RevForm.vaSpread1.RowHeight(-1) = 22
      Else
        COne = -0.5 '6
        coladj = 0 '2.3
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread1.RowHeight(0) = 18
        RevForm.vaSpread1.RowHeight(-1) = 16
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 8
        coladj = 6
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 12
          Next y
        Next x
      Else
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread1.RowHeight(0) = 17
'        RevForm.vaSpread1.FontBold = True
        RevForm.vaSpread1.RowHeight(-1) = 16
        COne = 0.5
        coladj = 0
      End If
      Case 800
        COne = 0 '-0.6
        coladj = 0 '1.55
        For x = 0 To 8
          For y = 0 To 2
            RevForm.vaSpread1.FontName = "Tahoma"
            RevForm.vaSpread1.Col = y
            RevForm.vaSpread1.Row = x
            RevForm.vaSpread1.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread1.RowHeight(0) = 14
        RevForm.vaSpread1.RowHeight(-1) = 14.75
      Case Else
       
    End Select
    RevForm.vaSpread1.ColWidth(1) = RevForm.vaSpread1.ColWidth(1) + COne
    RevForm.vaSpread1.ColWidth(2) = RevForm.vaSpread1.ColWidth(2) + coladj
    RevForm.vaSpread1.ColWidth(3) = 0
    RevForm.vaSpread1.ColWidth(4) = 0

    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 21
        coladj = 10
        For x = 0 To 10
          For y = 0 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 16
          Next y
        Next x
        RevForm.vaSpread2.RowHeight(-1) = 27.5
        RevForm.vaSpread2.RowHeight(0) = 27.5
      Else
        COne = -2 '11.25
        coladj = 1 '3.45
        For x = 0 To 10
          For y = 1 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 10 '12
          Next y
        Next x
        RevForm.vaSpread2.Col = 3
        RevForm.vaSpread2.Row = 0
        RevForm.vaSpread2.FontName = "Tahoma"
        RevForm.vaSpread2.FontSize = 10 '12
        RevForm.vaSpread2.Text = RevForm.vaSpread2.Text
        RevForm.vaSpread2.RowHeight(-1) = 18 '23.5
        RevForm.vaSpread2.RowHeight(0) = 18 '23.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 15
        coladj = 7
        For x = 0 To 10
          For y = 0 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 14
          Next y
        Next x
        RevForm.vaSpread2.RowHeight(0) = 24
        RevForm.vaSpread2.RowHeight(-1) = 22
      Else
        COne = 1 '6
        coladj = 0 '2.3
        For x = 0 To 10
          For y = 0 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread2.RowHeight(0) = 15.5
        RevForm.vaSpread2.RowHeight(-1) = 15.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 7
        coladj = 6
        For x = 0 To 10
          For y = 0 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 12
          Next y
        Next x
        RevForm.vaSpread2.RowHeight(0) = 19.5
'        RevForm.vaSpread2.FontBold = True
        RevForm.vaSpread2.RowHeight(-1) = 19.5
      Else
        COne = 1
        coladj = 0.5
      End If
      Case 800
        vaTabPro1.Font = "Tahoma"
        vaTabPro1.FontSize = 8
        COne = 0 '-0.6
        coladj = 0.5  '1.55
        For x = 0 To 10
          For y = 1 To 3
            RevForm.vaSpread2.FontName = "Tahoma"
            RevForm.vaSpread2.Col = y
            RevForm.vaSpread2.Row = x
            RevForm.vaSpread2.FontSize = 10
          Next y
        Next x
        RevForm.vaSpread2.RowHeight(0) = 14
        RevForm.vaSpread2.RowHeight(-1) = 14.75
      Case Else
       
    End Select
    RevForm.vaSpread2.ColWidth(1) = RevForm.vaSpread2.ColWidth(1) + COne
    RevForm.vaSpread2.ColWidth(2) = RevForm.vaSpread2.ColWidth(2) + coladj
    RevForm.vaSpread2.ColWidth(3) = RevForm.vaSpread2.ColWidth(2) + coladj
    RevForm.vaSpread2.ColWidth(4) = 0


End Function

Private Function CheckPersPayOrder() As Boolean
  Dim x As Integer, y As Integer
  Dim ThisOrder As Integer
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  Opt1Desc$ = ""
  Opt2Desc$ = ""
  Opt3Desc$ = ""
  CheckPersPayOrder = True 'no problem yet
  RevForm.vaSpread2.Col = 1
  RevForm.vaSpread2.Row = 8
  If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
    Opt1 = True
    Opt1Desc$ = QPTrim$(RevForm.vaSpread2.Text)
  End If
  RevForm.vaSpread2.Row = 9
  If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
    Opt2 = True
    Opt2Desc$ = QPTrim$(RevForm.vaSpread2.Text)
  End If
  RevForm.vaSpread2.Row = 10
  If QPTrim$(RevForm.vaSpread2.Text) <> "" Then
    Opt3 = True
    Opt3Desc$ = QPTrim$(RevForm.vaSpread2.Text)
  End If
  
  RevForm.vaSpread2.Col = 3
  For x = 1 To 10
    RevForm.vaSpread2.Row = x
    ThisOrder = CInt(RevForm.vaSpread2.Text)
    For y = 1 To 10
      RevForm.vaSpread2.Row = y
      If y = x Then GoTo SkipIt
      If ThisOrder = CInt(RevForm.vaSpread2.Text) Then
        CheckPersPayOrder = False
        Call TaxMsg(800, "The personal pay order for row " + CStr(y) + " and row " + CStr(x) + " are the same (" + CStr(ThisOrder) + "). Please reorder the personal revenues so that there are no duplicates.")
        Exit For
      End If
      If y <= 10 Then
        Exit Function
      End If
SkipIt:
    Next y
  Next x
  
End Function

Private Sub PrintMdltwnReal()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim TownName$, Add1$, Add2$, Add3$, Add4$
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  
  DueDate$ = Date
  RealTaxRate# = 0.25
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  Add4$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  If TownName = "" Then
    TownName = "Your Town"
  End If
  
  If Add1$ = "" Then
    Add1$ = "100 Main St"
  End If
  
  If QPTrim$(TaxMasterRec.City) = "" Then
    Add3$ = "Your Town, NC 27330"
    Add4$ = "Your Town, NC 27330"
  End If
  
  File$ = StartPath$ + "/TxBMdltwnRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  GoSub LoadHeaders
  
  Tab1 = 44 - Len(TownName) / 2
  Tab2 = 44 - Len(Add1) / 2
  Tab3 = 44 - Len(Add3) / 2

  Print #RptFile, "                                R E A L   E S T A T E"
  Print #RptFile, "                                 T A X   N O T I C E"
  Print #RptFile, Tab(Tab1); TownName
  Print #RptFile, Tab(Tab2); Add1
  Print #RptFile, Tab(Tab3); Add3
  Print #RptFile,
  Print #RptFile, "            VALUATION AMOUNT: "; Using$("$##,###,###.00", 150000);
  Print #RptFile, Tab(50); "ACCT. #: "; "125"
  Print #RptFile, "                   EXEMPTION: "; Using$("$##,###,###.00", 0);
  Print #RptFile, Tab(50); "PIN. #: "; "123456"
  Print #RptFile, "         LATE PENALTY AMOUNT: "; Using$("$##,###,###.00", 0);
  Print #RptFile, Tab(50); "RECPT #: "; Using$("#####0", 34)
  Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", 375);
  Print #RptFile, Tab(50); "TAX RATE %: "; Using$("#0.0000", RealTaxRate#)
  Print #RptFile, Tab(50); "TAX YEAR: "; Mid(Date, 6, 4)
  Print #RptFile, Tab(50); "DUE DATE: "; DueDate$
  Print #RptFile, Tab(11); Left$("JOHN PUBLIC", 45)
  Print #RptFile, Tab(11); Left$("1000 MAPLE ST", 35)
  Print #RptFile, Tab(11); Left$("PO BOX 120", 35)
  Print #RptFile, Tab(11); Add4$
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, Tab(31); "T H A N K   Y O U ! ! !"
  Print #RptFile,
  Print #RptFile, "~"
  Close
  ViewPrint File$, "Real Property Tax Bills", True
  Exit Sub
  
LoadHeaders:
  WorkName = ""
  WordLen = Len(TownName)
  For y = 1 To WordLen
    ch = Mid(TownName, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  TownName = WorkName
  
  WorkName = ""
  WordLen = Len(Add1)
  For y = 1 To WordLen
    ch = Mid(Add1, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add1 = WorkName
  
  WorkName = ""
  WordLen = Len(Add2)
  For y = 1 To WordLen
    ch = Mid(Add2, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add2 = WorkName
  
  WorkName = ""
  WordLen = Len(Add3)
  For y = 1 To WordLen
    ch = Mid(Add3, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add3 = WorkName
  
  Return

End Sub

Private Sub PrintMdltwnPers()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long, PersTaxRate#
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim TownName$, Add1$, Add2$, Add3$, Add4$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc$, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim Tab1 As Integer, Tab2 As Integer, Tab3 As Integer, Tab4 As Integer
  Dim DueDate$, WorkName$
  Dim FBill&
  Dim LBill&
  
  DueDate$ = Date
  PersTaxRate# = 0.25
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  Add4$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  If TownName = "" Then
    TownName = "Your Town"
  End If
  
  If Add1$ = "" Then
    Add1$ = "100 Main St"
  End If
  
  If QPTrim$(TaxMasterRec.City) = "" Then
    Add3$ = "Your Town, NC 27330"
    Add4$ = "Your Town, NC 27330"
  End If
  File$ = StartPath$ + "/TxBStandPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  Tab1 = 40 - Len(TownName) / 2
  Tab2 = 40 - Len(Add1) / 2
  Tab3 = 40 - Len(Add3) / 2
  
  CustName$ = "JOHN PUBLIC"
  
  Print #RptFile,
  Print #RptFile, Tab(Tab1); TownName
  Print #RptFile, Tab(Tab2); Add1
  Print #RptFile, Tab(Tab3); Add3
  Print #RptFile, Tab(27); "PERSONAL PROPERTY TAX BILL"
  Print #RptFile, Tab(30);
  For LC = 6 To 8
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(10); "ACCT # "; Using$("######0", 325);
  Print #RptFile, Tab(63); "BILL # "; Using$("######0", 17)
  Print #RptFile, Tab(10); Left$(CustName$, 25);
  Print #RptFile, Tab(63); "TAX YEAR: "; Mid(Date, 6, 4)
  Print #RptFile, Tab(10); Left$("100 MAPLE ST", 25);
  Print #RptFile, Tab(63); "TAX RATE: "; Using("#0.##0", PersTaxRate#) + "%"
  Print #RptFile, Tab(10); Left$("PO BOX 100", 25)
  Print #RptFile, Tab(10); Add3$
  For LC = 14 To 17
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(37); "PROPERTY"; Tab(51); "   TAX"; Tab(61); "   PPTRA"
  Print #RptFile, Tab(38); "  VALUE"; Tab(51); "AMOUNT"; Tab(61); "DISCOUNT"; Tab(71); "TOTAL DUE"
  'Line 23 Starts Here
  Print #RptFile, Tab(2); "Personal Property";
  Print #RptFile, Tab(38); Using$("###,##0", 26450);
  Print #RptFile, Tab(47); Using$("###,##0.00", 66.12);
  Print #RptFile, Tab(59); Using("###,##0.00", 46.96);
  Print #RptFile, Tab(70); Using("###,##0.00", 19.16)

  Print #RptFile,
  Print #RptFile, " PPTRA Information"

  Print #RptFile, Tab(2); "*" + "VIN# 1G8AHY67145Z30167";
  Print #RptFile, Tab(38); Using("###,##0", 8225);
  Print #RptFile, Tab(47); Using("###,##0.00", 20.56);
  Print #RptFile, Tab(59); Using("###,##0.00", 14.4)
  
  Print #RptFile, Tab(2); "*" + "VIN# 13255987411";
  Print #RptFile, Tab(38); Using("###,##0", 900);
  Print #RptFile, Tab(47); Using("###,##0.00", 2.25);
  Print #RptFile, Tab(59); Using("###,##0.00", 2.25)
  
  Print #RptFile, Tab(2); "*" + "VIN# 2L57PU669886711";
  Print #RptFile, Tab(38); Using("###,##0", 8200);
  Print #RptFile, Tab(47); Using("###,##0.00", 20.5);
  Print #RptFile, Tab(59); Using("###,##0.00", 14.35)
  
  Print #RptFile, Tab(2); "*" + "VIN# FTU9P4P99HY0678";
  Print #RptFile, Tab(38); Using("###,##0", 9125);
  Print #RptFile, Tab(47); Using("###,##0.00", 22.81);
  Print #RptFile, Tab(59); Using("###,##0.00", 15.97)
  
  ' Finish the bill up here
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, Tab(40); "Total Tax Due by "; DueDate$;
  'Put Late Here and Add to Total
  Print #RptFile, Tab(69); Using("$###,##0.00", 19.16)
   
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, " The Personal Property Relief Act provides that the tax on the first"
  Print #RptFile, " " + Using$("$##,##0.00", TaxMasterRec.MaxVehTaxVal) + " of value of your personal car, motorcycle and pickup  "
  Print #RptFile, " or panel truck under 7,501 pounds, which is a qualifying vehicle, has been"
  Print #RptFile, " reduced by " + Using$("#0.00", TaxMasterRec.PPTRADisc) + "% this year. If your qualifying vehicle's value is"
  Print #RptFile, " " + Using$("$#,##0.00", TaxMasterRec.MinVehTaxVal) + " or less, your tax has been eliminated. These reductions are"
  Print #RptFile, " based on the local tax rates in effect on July 1 or August 1, 1997,"
  Print #RptFile, " whichever was higher. Please contact the Town Office with any questions."
  Print #RptFile, ""
  Print #RptFile, ""

  Print #RptFile,
  Print #RptFile,
  Print #RptFile,
  
  Close
  ViewPrint File$, "Pers Property Tax Bills", True

End Sub

Private Sub PrintCdrBluffPers()
  Dim PersTaxRate#
  Dim PYear As Integer
  Dim File$
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim TownName$, Add1$, Add2$, Add3$
  Dim VehDesc$, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim DueDate$, VehDisc$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisDate As Integer
  
  ThisDate = Date2Num(Date)
  ThisDate = ThisDate + 60
  DueDate = MakeRegDate(ThisDate)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If TownName = "" Then
    TownName = "Your Town"
  End If
  
  If Add1$ = "" Then
    Add1$ = "100 Main St"
  End If
  
  If QPTrim$(TaxMasterRec.City) = "" Then
    Add3$ = "Your Town, NC 27330"
  End If
  
  File$ = StartPath$ + "/TxCdrBluffPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  PersTaxRate# = 0.25
  
  CustName$ = "Joe Smith"
  Print #RptFile, "~"; Tab(34); Mid(DueDate, 7, 4); " PERSONAL PROPERTY"
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, Tab(5); "   "
  
  Print #RptFile, Tab(10); "111-22-3333"; Tab(65); "PP"; Using("##.###", PersTaxRate#)
  Print #RptFile, " "
  Print #RptFile, " "

  PYear = CInt(Mid(Date, 7, 4))
  VehDisc$ = "2005 Ford F-150"
  Print #RptFile, VehDesc$;
  Print #RptFile, Tab(33); Using("##,###,###", 21000);
  Print #RptFile, Tab(44); Using("###,###.##", 52.5);
  Print #RptFile, Tab(54); Using("#####.##", 35);
  Print #RptFile, Tab(64); Using("##,###.##", 17.5)

  VehDisc$ = "2001 Chevy Impala"
  Print #RptFile, VehDesc$;
  Print #RptFile, Tab(33); Using("##,###,###", 8000);
  Print #RptFile, Tab(44); Using("###,###.##", 20);
  Print #RptFile, Tab(54); Using("#####.##", 14);
  Print #RptFile, Tab(64); Using("##,###.##", 6)
  
  For LCnt = 3 To 5
    Print #RptFile, ""
  Next
  Print #RptFile, ""
  Print #RptFile, Tab(10); Using("#####", 1);
  Print #RptFile, Tab(36); DueDate$; Tab(62); Using("$###,###.##", 23.5)
  Print #RptFile,
  Print #RptFile, Tab(9); CustName$
  Print #RptFile, Tab(9); "1234 Elm Street"
  Print #RptFile, Tab(9); "PO Box 1234"
  Print #RptFile, Tab(9); "Anytown, VA 24609"
  Print #RptFile,
  Print #RptFile,
  Print #RptFile, "~"
  
  Close
  ViewPrint File$, "Personal Property Tax Bills", True

End Sub

Private Sub PrintCdrBluffReal()
  Dim x As Long, RealTaxRate#
  Dim File$
  Dim CustName As String * 45
  Dim RptFile#
  Dim TownName$, Add1$, Add2$, Add3$
  Dim WhatYear As Integer
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, ThisDate As Integer
  
  ThisDate = Date2Num(Date)
  ThisDate = ThisDate + 60
  DueDate = MakeRegDate(ThisDate)
  
  RealTaxRate# = 0.25
  File$ = StartPath$ + "/TxBCdrBluffRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile

  CustName$ = "Joe Smith"
  WhatYear = CInt(Mid(Date, 7, 4))
  Print #RptFile, "~"
  Print #RptFile, Tab(50); CStr(WhatYear); Tab(78); Using("########", 1234)
  Print #RptFile,
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(28); Using("#.##", 0.25);
  Print #RptFile, Tab(36); Using("###,###,###", 40000);
  Print #RptFile, Tab(48); Using("##,###,###", 150000);
  Print #RptFile, Tab(61); Using("##,###,###", 190000);
  Print #RptFile, Tab(75); Using("###,###.##", 475);
  Print #RptFile, Tab(90); "5%"
  Print #RptFile, " "
  Print #RptFile, Tab(68); DueDate$
  Print #RptFile, "Map #2N6 Block #4R"
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, ""
  Print #RptFile, Tab(7); "ACCT # "; "1001"
  Print #RptFile, Tab(7); Left$(CustName$, 45)
  Print #RptFile, Tab(7); Left$("1234 Elm Street", 35)
  Print #RptFile, Tab(7); Left$("PO Box 1234", 35)
  Print #RptFile, Tab(7); "Anytown, VA 24609"
  Print #RptFile,
  Print #RptFile, "BN"; Using("#####", 1)
  Print #RptFile, "~"

  Close
  
  ViewPrint File$, "Real Property Tax Bills", True

End Sub












