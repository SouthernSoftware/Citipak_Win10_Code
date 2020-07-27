VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEditEmpData 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmNewEmpData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6180
      Left            =   384
      TabIndex        =   102
      Top             =   1746
      Width           =   10956
      _Version        =   196609
      _ExtentX        =   19325
      _ExtentY        =   10901
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
      TabCount        =   8
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
      TabCaption      =   "frmNewEmpData.frx":08CA
      PageEarMarkPictureNext=   "frmNewEmpData.frx":0CE5
      PageEarMarkPicturePrev=   "frmNewEmpData.frx":0D01
      EarMarkPictureNext=   "frmNewEmpData.frx":0D1D
      EarMarkPicturePrev=   "frmNewEmpData.frx":0D39
      Begin ImpproLib.vaImprint vaImprint6 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   153
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":0D55
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   1
            Left            =   2880
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   96
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   20
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
            TabIndex        =   22
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
            TabIndex        =   24
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
            TabIndex        =   26
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
            TabIndex        =   19
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
            TabIndex        =   21
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
            TabIndex        =   23
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
            TabIndex        =   25
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
         Begin VB.Shape Shape8 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
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
            TabIndex        =   156
            Top             =   1104
            Width           =   972
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
            TabIndex        =   155
            Top             =   1104
            Width           =   1884
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
            TabIndex        =   154
            Top             =   1104
            Width           =   1308
         End
      End
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   136
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":0D71
         Begin LpLib.fpCombo fpcomboEIC 
            Height          =   405
            Left            =   4710
            TabIndex        =   41
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
            ColDesigner     =   "frmNewEmpData.frx":0D8D
         End
         Begin LpLib.fpCombo fpcomboMedX 
            Height          =   405
            Left            =   6000
            TabIndex        =   40
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
            ColDesigner     =   "frmNewEmpData.frx":10DC
         End
         Begin LpLib.fpCombo fpcomboSocX 
            Height          =   405
            Left            =   6330
            TabIndex        =   39
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
            ColDesigner     =   "frmNewEmpData.frx":13D3
         End
         Begin LpLib.fpCombo fpcomboFedStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   30
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
            ColDesigner     =   "frmNewEmpData.frx":16CA
         End
         Begin LpLib.fpCombo fpcomboStateStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   36
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
            ColDesigner     =   "frmNewEmpData.frx":1A19
         End
         Begin LpLib.fpCombo fpcomboStateAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   34
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
            ColDesigner     =   "frmNewEmpData.frx":1D68
         End
         Begin LpLib.fpCombo fpcomboFedAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   28
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
            ColDesigner     =   "frmNewEmpData.frx":205F
         End
         Begin LpLib.fpCombo fpcomboStateX 
            Height          =   405
            Left            =   1935
            TabIndex        =   33
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
            ColDesigner     =   "frmNewEmpData.frx":2356
         End
         Begin LpLib.fpCombo fpcomboFedX 
            Height          =   405
            Left            =   1935
            TabIndex        =   27
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
            ColDesigner     =   "frmNewEmpData.frx":264D
         End
         Begin EditLib.fpCurrency fptxtAddWHFed 
            Height          =   396
            Left            =   8352
            TabIndex        =   32
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
            TabIndex        =   37
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
            TabIndex        =   31
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
            TabIndex        =   38
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
            TabIndex        =   29
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
            TabIndex        =   35
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
            TabIndex        =   150
            Top             =   1056
            Width           =   1020
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
            TabIndex        =   148
            Top             =   3936
            Width           =   1212
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
            TabIndex        =   147
            Top             =   3360
            Width           =   2124
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
            TabIndex        =   146
            Top             =   2784
            Width           =   2700
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
            TabIndex        =   145
            Top             =   768
            Width           =   1188
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
            TabIndex        =   144
            Top             =   1056
            Width           =   1356
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
            TabIndex        =   143
            Top             =   1056
            Width           =   828
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
            TabIndex        =   142
            Top             =   1056
            Width           =   732
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
            TabIndex        =   141
            Top             =   1056
            Width           =   972
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
            TabIndex        =   140
            Top             =   768
            Width           =   636
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
            TabIndex        =   139
            Top             =   1056
            Width           =   1068
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
            TabIndex        =   138
            Top             =   2208
            Width           =   684
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
            TabIndex        =   137
            Top             =   1680
            Width           =   828
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
      End
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   124
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":2944
         Begin LpLib.fpCombo fpcomboPayType 
            Height          =   405
            Left            =   2685
            TabIndex        =   62
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
            ColDesigner     =   "frmNewEmpData.frx":2960
         End
         Begin LpLib.fpCombo fpcomboStatus 
            Height          =   405
            Left            =   2685
            TabIndex        =   61
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
            ColDesigner     =   "frmNewEmpData.frx":2C57
         End
         Begin LpLib.fpCombo fpcomboFreq 
            Height          =   405
            Left            =   2685
            TabIndex        =   63
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
            ColDesigner     =   "frmNewEmpData.frx":2F4E
         End
         Begin EditLib.fpCurrency fptxtRate 
            Height          =   450
            Left            =   7830
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   59
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
            TabIndex        =   60
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   69
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
            TabIndex        =   66
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
            TabIndex        =   70
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
            TabIndex        =   191
            Top             =   4320
            Width           =   1185
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
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
            TabIndex        =   135
            Top             =   3600
            Width           =   2070
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
            TabIndex        =   134
            Top             =   3030
            Width           =   1500
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
            TabIndex        =   133
            Top             =   2400
            Width           =   1215
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
            TabIndex        =   132
            Top             =   1830
            Width           =   1215
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
            TabIndex        =   131
            Top             =   1290
            Width           =   870
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
            TabIndex        =   130
            Top             =   3600
            Width           =   1500
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
            TabIndex        =   129
            Top             =   2970
            Width           =   1455
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
            TabIndex        =   128
            Top             =   2400
            Width           =   1215
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
            TabIndex        =   127
            Top             =   1830
            Width           =   1065
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
            TabIndex        =   126
            Top             =   1245
            Width           =   1260
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
            TabIndex        =   125
            Top             =   645
            Width           =   690
         End
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   5100
         Left            =   255
         TabIndex        =   103
         Top             =   765
         Width           =   10230
         _Version        =   196609
         _ExtentX        =   18045
         _ExtentY        =   8996
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":3245
         Begin LpLib.fpCombo fpcomboRetType 
            Height          =   405
            Left            =   7680
            TabIndex        =   90
            ToolTipText     =   "Select the Employee's Gender from the pick list."
            Top             =   2970
            Width           =   1770
            _Version        =   196608
            _ExtentX        =   3122
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
            Columns         =   1
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
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmNewEmpData.frx":3261
         End
         Begin LpLib.fpCombo fpcomboGender 
            Height          =   405
            Left            =   7680
            TabIndex        =   87
            ToolTipText     =   "Select the Employee's Gender from the pick list."
            Top             =   1200
            Width           =   1770
            _Version        =   196608
            _ExtentX        =   3122
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
            Columns         =   1
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
            ColDesigner     =   "frmNewEmpData.frx":3584
         End
         Begin VB.CheckBox chkRet 
            BackColor       =   &H008F8265&
            Caption         =   "Retired?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   8400
            TabIndex        =   190
            ToolTipText     =   "Click to exclude this employee from being retirement matched. This employee will appear on state retirement reports."
            Top             =   2520
            Width           =   1095
         End
         Begin VB.CheckBox chkTemp 
            BackColor       =   &H008F8265&
            Caption         =   "Temporary?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   6360
            TabIndex        =   189
            ToolTipText     =   $"frmNewEmpData.frx":38A7
            Top             =   2520
            Width           =   1335
         End
         Begin EditLib.fpMask txtZip 
            Height          =   396
            Left            =   3648
            TabIndex        =   84
            Top             =   2928
            Width           =   1788
            _Version        =   196608
            _ExtentX        =   3154
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
         Begin EditLib.fpText fptxtRetNum 
            Height          =   396
            Left            =   7680
            TabIndex        =   89
            ToolTipText     =   "Enter the Employee's Retirement System Number here."
            Top             =   2064
            Width           =   1752
            _Version        =   196608
            _ExtentX        =   3090
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
            MaxLength       =   16
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
         Begin EditLib.fpMask fpMaskSoc 
            Height          =   396
            Left            =   7680
            TabIndex        =   85
            ToolTipText     =   "Enter the Employee's Social Security Number here."
            Top             =   336
            Width           =   1752
            _Version        =   196608
            _ExtentX        =   3087
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
            Mask            =   "###-##-####"
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
         Begin EditLib.fpText txtState 
            Height          =   396
            Left            =   2112
            TabIndex        =   83
            ToolTipText     =   "Enter the Employee's State of Residence here."
            Top             =   2928
            Width           =   732
            _Version        =   196608
            _ExtentX        =   1291
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
            CharValidationText=   "A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, ., , "
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
         Begin EditLib.fpText txtCity 
            Height          =   396
            Left            =   2112
            TabIndex        =   82
            ToolTipText     =   "Enter the Employee's City of Residence here."
            Top             =   2496
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            MaxLength       =   24
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
         Begin EditLib.fpText txtAddress2 
            Height          =   396
            Left            =   2112
            TabIndex        =   81
            ToolTipText     =   "Enter the Employee's Address here."
            Top             =   2064
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            MaxLength       =   36
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
         Begin EditLib.fpText txtAddress1 
            Height          =   396
            Left            =   2112
            TabIndex        =   80
            ToolTipText     =   "Enter the Employee's Address here."
            Top             =   1632
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            MaxLength       =   36
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
         Begin EditLib.fpText txtFirstName 
            Height          =   396
            Left            =   2112
            TabIndex        =   79
            ToolTipText     =   "Enter the Employee's First Name here."
            Top             =   1200
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            MaxLength       =   24
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
         Begin EditLib.fpText txtNumber 
            Height          =   396
            Left            =   2112
            TabIndex        =   77
            ToolTipText     =   "Assign an Employee Number here."
            Top             =   336
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            CharValidationText=   "1,2,3,4,5,6,7,8,9,0"
            MaxLength       =   10
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
         Begin EditLib.fpText txtLastName 
            Height          =   396
            Left            =   2112
            TabIndex        =   78
            ToolTipText     =   "Enter the Employee's Last Name here."
            Top             =   768
            Width           =   3324
            _Version        =   196608
            _ExtentX        =   5863
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
            MaxLength       =   24
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
         Begin EditLib.fpDateTime fpMaskBDay 
            Height          =   372
            Left            =   7680
            TabIndex        =   86
            ToolTipText     =   "Enter the Employee's Date of Birth here."
            Top             =   768
            Width           =   1776
            _Version        =   196608
            _ExtentX        =   3136
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
         Begin EditLib.fpText fptxtRace 
            Height          =   396
            Left            =   7680
            TabIndex        =   88
            ToolTipText     =   "Enter the Employee's race here."
            Top             =   1632
            Width           =   1752
            _Version        =   196608
            _ExtentX        =   3090
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
            CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z a b c d e f g h i j c l m n o p q r s t u v w x y z"
            MaxLength       =   14
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
         Begin EditLib.fpText fptxtMainDept 
            Height          =   390
            Left            =   7680
            TabIndex        =   91
            ToolTipText     =   "Enter the department from which most of this employee's pay is allocatted."
            Top             =   3405
            Width           =   1755
            _Version        =   196608
            _ExtentX        =   3090
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   20
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
         Begin EditLib.fpText fptxtContactName 
            Height          =   390
            Left            =   2355
            TabIndex        =   92
            ToolTipText     =   "Enter the emergency contact's name."
            Top             =   4050
            Width           =   2805
            _Version        =   196608
            _ExtentX        =   4953
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
            CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z a b c d e f g h i j c l m n o p q r s t u v w x y z"
            MaxLength       =   48
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
         Begin EditLib.fpText fptxtRelationship 
            Height          =   390
            Left            =   2355
            TabIndex        =   93
            ToolTipText     =   "Enter the emergency contact's relationship (wife, brother, etc.)."
            Top             =   4485
            Width           =   2805
            _Version        =   196608
            _ExtentX        =   4953
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
            CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z a b c d e f g h i j c l m n o p q r s t u v w x y z"
            MaxLength       =   24
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
         Begin EditLib.fpMask fptxtHomePhone 
            Height          =   390
            Left            =   7245
            TabIndex        =   95
            ToolTipText     =   "Enter the Employee's home phone number here."
            Top             =   4485
            Width           =   2190
            _Version        =   196608
            _ExtentX        =   3873
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
         Begin EditLib.fpMask fptxtContactPhone 
            Height          =   390
            Left            =   7245
            TabIndex        =   94
            ToolTipText     =   "Enter the emergency contact's phone number."
            Top             =   4050
            Width           =   2190
            _Version        =   196608
            _ExtentX        =   3873
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
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship"
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
            Left            =   720
            TabIndex        =   188
            Top             =   4635
            Width           =   1545
         End
         Begin VB.Label Label75 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact  Phone"
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
            Left            =   5430
            TabIndex        =   187
            Top             =   4200
            Width           =   1740
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Name"
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
            Height          =   255
            Left            =   720
            TabIndex        =   186
            Top             =   4200
            Width           =   1650
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone"
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
            Left            =   5430
            TabIndex        =   185
            Top             =   4635
            Width           =   1410
         End
         Begin VB.Label Label72 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Emergency Contact Information"
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
            Height          =   300
            Left            =   360
            TabIndex        =   184
            Top             =   3680
            Width           =   3465
         End
         Begin VB.Label Label71 
            BackStyle       =   0  'Transparent
            Caption         =   "Main Dept."
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
            Left            =   5910
            TabIndex        =   183
            Top             =   3525
            Width           =   1500
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Race"
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
            Left            =   5904
            TabIndex        =   181
            Top             =   1776
            Width           =   1500
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "BirthDate"
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
            Left            =   5904
            TabIndex        =   117
            Top             =   912
            Width           =   1596
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Ret Type"
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
            Left            =   5910
            TabIndex        =   116
            Top             =   3090
            Width           =   1545
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Ret Number"
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
            Left            =   5904
            TabIndex        =   115
            Top             =   2188
            Width           =   1500
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender*"
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
            Left            =   5904
            TabIndex        =   114
            Top             =   1344
            Width           =   1500
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Soc Sec Num*"
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
            Left            =   5904
            TabIndex        =   113
            Top             =   480
            Width           =   1644
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip*"
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
            Left            =   3120
            TabIndex        =   112
            Top             =   3024
            Width           =   444
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "State*"
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
            Left            =   720
            TabIndex        =   111
            Top             =   3024
            Width           =   1164
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "City*"
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
            Left            =   672
            TabIndex        =   110
            Top             =   2640
            Width           =   1020
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2"
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
            Left            =   672
            TabIndex        =   109
            Top             =   2208
            Width           =   1164
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1*"
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
            TabIndex        =   108
            Top             =   1728
            Width           =   1308
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4770
            Left            =   120
            Top             =   195
            Width           =   9945
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name*"
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
            TabIndex        =   106
            Top             =   1296
            Width           =   1356
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Number*"
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
            TabIndex        =   105
            Top             =   480
            Width           =   1404
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name*"
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
            Left            =   672
            TabIndex        =   104
            Top             =   864
            Width           =   1644
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   10050
            Y1              =   3915
            Y2              =   3927
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   107
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":3937
         Begin LpLib.fpCombo fpcomboPrenoted 
            Height          =   405
            Left            =   2685
            TabIndex        =   73
            ToolTipText     =   "Enter a ""Y"" if this employee has been prenoted."
            Top             =   2160
            Width           =   2025
            _Version        =   196608
            _ExtentX        =   3572
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
            ColDesigner     =   "frmNewEmpData.frx":3953
         End
         Begin LpLib.fpCombo fpcomboBankdraft 
            Height          =   405
            Left            =   2685
            TabIndex        =   71
            ToolTipText     =   "No help for this field."
            Top             =   1440
            Width           =   2025
            _Version        =   196608
            _ExtentX        =   3572
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
            ColDesigner     =   "frmNewEmpData.frx":3C4A
         End
         Begin EditLib.fpText txtBankTransNo 
            Height          =   396
            Left            =   7440
            TabIndex        =   76
            ToolTipText     =   "No help for this field."
            Top             =   2832
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
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
            MaxLength       =   9
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
         Begin EditLib.fpText txtBankName 
            Height          =   396
            Left            =   7440
            TabIndex        =   74
            ToolTipText     =   "No help for this field."
            Top             =   2112
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
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
            MaxLength       =   33
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
         Begin EditLib.fpText txtBankAcctNo 
            Height          =   396
            Left            =   7440
            TabIndex        =   72
            ToolTipText     =   "Enter the Employee's Direct Deposit Account Number here."
            Top             =   1392
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
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
            MaxLength       =   20
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
         Begin EditLib.fpText txtBankLocation 
            Height          =   396
            Left            =   2688
            TabIndex        =   75
            ToolTipText     =   "No help for this field."
            Top             =   2832
            Width           =   2028
            _Version        =   196608
            _ExtentX        =   3577
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
            MaxLength       =   30
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
         Begin VB.Label LabelDD 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmNewEmpData.frx":3F99
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
            Height          =   735
            Left            =   1140
            TabIndex        =   192
            Top             =   360
            Width           =   8175
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Transit No"
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
            Height          =   444
            Left            =   5472
            TabIndex        =   123
            Top             =   2928
            Width           =   1836
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Left            =   5472
            TabIndex        =   122
            Top             =   2256
            Width           =   1548
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Acct No"
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
            Left            =   5472
            TabIndex        =   121
            Top             =   1536
            Width           =   1788
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Location"
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
            Left            =   864
            TabIndex        =   120
            Top             =   2976
            Width           =   1788
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Prenoted"
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
            Left            =   864
            TabIndex        =   119
            Top             =   2304
            Width           =   1068
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "BankDraft Code"
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
            Left            =   864
            TabIndex        =   118
            Top             =   1584
            Width           =   1836
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
      End
      Begin ImpproLib.vaImprint vaImprint5 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   151
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":40A1
         Begin FPSpread.vaSpread vaSpreadMisc 
            Height          =   4260
            Left            =   360
            TabIndex        =   152
            Top             =   360
            Width           =   9525
            _Version        =   196613
            _ExtentX        =   16806
            _ExtentY        =   7514
            _StockProps     =   64
            ColsFrozen      =   4
            EditEnterAction =   5
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   13684944
            MaxCols         =   4
            MaxRows         =   50
            MoveActiveOnFocus=   0   'False
            ProcessTab      =   -1  'True
            RetainSelBlock  =   0   'False
            ShadowColor     =   13684944
            SpreadDesigner  =   "frmNewEmpData.frx":40BD
            TextTip         =   2
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
      End
      Begin ImpproLib.vaImprint vaImprint7 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   157
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":4709
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   8
            Left            =   7776
            TabIndex        =   16
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
            TabIndex        =   14
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
            TabIndex        =   12
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
            TabIndex        =   10
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
            TabIndex        =   8
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
            TabIndex        =   6
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
            TabIndex        =   4
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
            TabIndex        =   2
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
            TabIndex        =   1
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
            TabIndex        =   3
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
            TabIndex        =   5
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
            TabIndex        =   7
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
            TabIndex        =   9
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
            TabIndex        =   11
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
            TabIndex        =   13
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
            TabIndex        =   15
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
            TabIndex        =   180
            Top             =   480
            Width           =   2220
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
            TabIndex        =   167
            Top             =   4224
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
            TabIndex        =   166
            Top             =   3792
            Width           =   252
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
            TabIndex        =   165
            Top             =   3312
            Width           =   300
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
            TabIndex        =   164
            Top             =   2832
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
            TabIndex        =   163
            Top             =   2352
            Width           =   348
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
            TabIndex        =   162
            Top             =   1872
            Width           =   300
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
            TabIndex        =   161
            Top             =   1344
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
            TabIndex        =   160
            Top             =   864
            Width           =   300
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
            TabIndex        =   159
            Top             =   240
            Width           =   2220
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
            TabIndex        =   158
            Top             =   384
            Width           =   1980
         End
      End
      Begin ImpproLib.vaImprint vaImprint8 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   168
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmNewEmpData.frx":4725
         Begin LpLib.fpCombo fpcombo401K 
            Height          =   405
            Left            =   5370
            TabIndex        =   57
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
            ColDesigner     =   "frmNewEmpData.frx":4741
         End
         Begin LpLib.fpCombo fpcomboESC 
            Height          =   405
            Left            =   9030
            TabIndex        =   58
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
            ColDesigner     =   "frmNewEmpData.frx":4A38
         End
         Begin LpLib.fpCombo fpcomboLT 
            Height          =   405
            Left            =   2070
            TabIndex        =   56
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
            ColDesigner     =   "frmNewEmpData.frx":4D2F
         End
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   2
            Left            =   3264
            TabIndex        =   47
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
            TabIndex        =   55
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   0
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
            TabIndex        =   42
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
            TabIndex        =   54
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
            TabIndex        =   52
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
            TabIndex        =   46
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
            TabIndex        =   53
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
            TabIndex        =   51
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
            TabIndex        =   45
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   50
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
            TabIndex        =   182
            Top             =   4272
            Width           =   1884
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   192
            X2              =   10128
            Y1              =   3840
            Y2              =   3840
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
            TabIndex        =   178
            Top             =   4272
            Width           =   2412
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
            TabIndex        =   177
            Top             =   4272
            Width           =   1500
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
            TabIndex        =   176
            Top             =   3264
            Width           =   924
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
            TabIndex        =   175
            Top             =   2688
            Width           =   972
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
            TabIndex        =   174
            Top             =   2112
            Width           =   1308
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
            TabIndex        =   173
            Top             =   1536
            Width           =   1164
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
            TabIndex        =   172
            Top             =   336
            Width           =   876
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
            TabIndex        =   171
            Top             =   384
            Width           =   540
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
            TabIndex        =   170
            Top             =   432
            Width           =   780
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
            TabIndex        =   169
            Top             =   960
            Width           =   972
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4668
            Left            =   192
            Top             =   192
            Width           =   9948
         End
      End
   End
   Begin EditLib.fpText fptxtHeader 
      Height          =   348
      Left            =   4080
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   1248
      Width           =   3420
      _Version        =   196608
      _ExtentX        =   6032
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   9720
      TabIndex        =   194
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   8056
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmNewEmpData.frx":5026
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   8100
      TabIndex        =   195
      TabStop         =   0   'False
      Top             =   8056
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmNewEmpData.frx":523A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdYTD 
      Height          =   495
      Left            =   6480
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   8056
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmNewEmpData.frx":544E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHistory 
      Height          =   495
      Left            =   4865
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   8056
      Width           =   1335
      _Version        =   131072
      _ExtentX        =   2355
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmNewEmpData.frx":5660
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   492
      Left            =   2911
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   8056
      Width           =   1668
      _Version        =   131072
      _ExtentX        =   2942
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmNewEmpData.frx":5873
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   360
      Left            =   9360
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1785
      _Version        =   131072
      _ExtentX        =   3149
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
      ButtonDesigner  =   "frmNewEmpData.frx":5A8B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMessage 
      Height          =   495
      Left            =   960
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   8056
      Width           =   1665
      _Version        =   131072
      _ExtentX        =   2937
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmNewEmpData.frx":5CA7
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Required Fields*"
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
      Height          =   348
      Left            =   480
      TabIndex        =   149
      Top             =   1344
      Width           =   1884
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2796
      TabIndex        =   101
      Top             =   456
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   972
      Index           =   1
      Left            =   1500
      Top             =   156
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1044
      Left            =   1500
      Top             =   96
      Width           =   8652
   End
End
Attribute VB_Name = "frmEditEmpData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim DedCnt As Integer
Public newEmpFlag As Boolean
Dim tempEmpNum As Long
Dim thisRecordNum As Integer 'assuming that a record number is passed in
'Dim NoCITIPAK As Boolean
Dim RetireFlag As Boolean
Dim SSN As String
Dim CurEmpNum As String
Dim BadGLNum As Boolean
Dim ListFlag As Boolean
Dim ExitFlag As Boolean
Dim BtnFnt As Double

Private Sub chkRet_Click()
  Dim thisNum$
  Dim ThisLen As Integer
  
  If chkRet.Value = 1 Then
    chkTemp.Value = 0
    If Mid(fptxtRetNum.Text, 1, 1) = "T" Then
      ThisLen = Len(QPTrim$(fptxtRetNum.Text))
      thisNum = Mid(fptxtRetNum.Text, 2, ThisLen)
      fptxtRetNum.Text = "R" + thisNum
    ElseIf Mid(fptxtRetNum.Text, 1, 1) = "R" Then
      Exit Sub
    Else
      fptxtRetNum.Text = "R" + fptxtRetNum.Text
    End If
  Else
    If Mid(fptxtRetNum.Text, 1, 1) = "R" Then
      ThisLen = Len(fptxtRetNum.Text)
      fptxtRetNum.Text = Mid(fptxtRetNum.Text, 2, ThisLen)
      fptxtRetNum.Text = QPTrim$(fptxtRetNum.Text)
    End If
  End If

End Sub
Private Sub chkTemp_Click()
  Dim thisNum$
  Dim ThisLen As Integer
  
  If chkTemp.Value = 1 Then
    chkRet.Value = 0
    If Mid(fptxtRetNum.Text, 1, 1) = "R" Then
      ThisLen = Len(fptxtRetNum.Text)
      thisNum = Mid(fptxtRetNum.Text, 2, ThisLen)
      fptxtRetNum.Text = "T" + thisNum
    ElseIf Mid(fptxtRetNum.Text, 1, 1) = "T" Then
      Exit Sub
    Else
      fptxtRetNum.Text = "T" + fptxtRetNum.Text
    End If
  Else
    If Mid(fptxtRetNum.Text, 1, 1) = "T" Then
      ThisLen = Len(fptxtRetNum.Text)
      fptxtRetNum.Text = Mid(fptxtRetNum.Text, 2, ThisLen)
      fptxtRetNum.Text = QPTrim$(fptxtRetNum.Text)
    End If
  End If

End Sub

Private Sub cmdExit_Click()
  'sends process to PR_Common module to examine every
  'field of all three Employee Maintenance forms
  'checking to see if any changes have been made before
  'exiting erroneously
  Dim EmpData2FileHandle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim RecLen As Long
   
  OpenEmpData2File EmpData2FileHandle
  RecLen = LOF(EmpData2FileHandle) \ Len(Emp2Rec)
  Close EmpData2FileHandle
  If RecLen = 0 Then
    If MsgBox("Do you want to exit without saving any changes?", vbYesNo) = vbYes Then
      frmEmployeeMaintMenu.Show
      DoEvents
      Unload frmEditEmpData
      Exit Sub
    Else
      Exit Sub
    End If
  End If
  Call checkExitEmp(newEmpFlag, thisRecordNum, Me)
  
End Sub

Private Sub cmdHistory_Click()
'this command button is hidden if this is a new employee
  frmTransHistModal.Show vbModal ', Me

End Sub

Private Sub cmdList_Click()
  cmdList.SetFocus
  frmGLPickList.Show vbModal, Me
  DoEvents
End Sub

Private Sub cmdMessage_Click()
  If RecNum > 0 Then
    frmPREmpMessage.Show vbModal
  End If
 
End Sub

Private Sub cmdPrint_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PrintText
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PrintGraphics
  Else
    Exit Sub
  End If
End Sub
Private Sub PrintText()
  Dim RptName As String, EDistAmt(1 To 8) As Double
  Dim Emp2RecLen As Integer, ECnt As Integer
  Dim DataFileSize As Long, cnt As Integer
  Dim DataNumOfRecs As Long, EDistAcct(1 To 8) As String
  Dim RptHandle As Integer
  Dim RptTitle As String
  Dim FF As String, x As Integer
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim TDate$, HDate$, BDate$
  Dim RDate$, Nextx As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  ReDim Emp2Data(1) As EmpData2Type
  ReDim Desc$(1 To 53)
  ReDim DedCodeRec(1 To 50) As DedCodeRecType
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim EmpZip$, ZipLen As Integer
  
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  InFileNames(4) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  
  MaxLines = 57
  FF$ = Chr$(12)
  RptName$ = "PRRPTS\EMPDATA.RPT"

  OpenDedCodeFile DedCodeFileHandle
  For cnt = 1 To 50
    Get DedCodeFileHandle, cnt, DedCodeRec(cnt)
    Desc$(cnt) = QPTrim$(DedCodeRec(cnt).DCDESC1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close DedCodeFileHandle
  OpenErnCodeFile ErnCodeFileHandle
  For cnt = 51 To 53
    Get ErnCodeFileHandle, cnt - 50, ErnCodeRec
    Desc$(cnt) = QPTrim$(ErnCodeRec.ERNCODE1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close ErnCodeFileHandle
 
  RptTitle$ = "Employee Information Listing"

  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  RPTSetupPRN 1, RptHandle
  OpenEmpData2File EmpData2FileHandle
  Get EmpData2FileHandle, thisRecordNum, EmpData2FileRec
  For ECnt = 1 To 8
     EDistAcct(ECnt) = EmpData2FileRec.EDist(ECnt).DAcct
     EDistAmt(ECnt) = EmpData2FileRec.EDist(ECnt).DAmt
  Next ECnt
  If Not EmpData2FileRec.Deleted Then
    GoSub PrintEmpData
  End If
  RPTSetupPRN 123, RptHandle '7/24 revised 8/15
  Close EmpData2FileHandle
  Close RptHandle

  ViewPrint RptName$, RptTitle$, True
  MainLog ("Employee Data File processed.")
  EnableCloseButton Me.hwnd, True

Exit Sub

PrintEmpData:
  EmpZip = QPTrim$(EmpData2FileRec.EmpZip) '06/08/04
  ZipLen = Len(EmpZip) '06/08/04
  If ZipLen > 5 Then
    EmpZip = Mid(EmpZip, 1, 5) + "-" + Mid(EmpZip, 6, ZipLen) '06/08/04
    EmpData2FileRec.EmpZip = EmpZip '06/08/04
  End If

  BDate$ = MakeRegDate(EmpData2FileRec.EMPBDAY)
  If BDate = "12/31/1979" Then BDate = "Not Saved  "
  HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
  If HDate = "12/31/1979" Then HDate = "Not Saved  "
  RDate = MakeRegDate(EmpData2FileRec.EMPRDATE)
  If RDate = "12/31/1979" Then RDate = "Not Saved  "
  TDate = MakeRegDate(EmpData2FileRec.EMPTDATE)
  If TDate = "12/31/1979" Then TDate = "Not Saved  "
  Print #RptHandle, "--------------------------------------------------------------------------------------------"
  Print #RptHandle, ""
  Print #RptHandle, "                  C O N F I D E N T I A L   D A T A   F I L E"
  Print #RptHandle, ""
  Print #RptHandle, "  Employee Information"
  Print #RptHandle, "       Number: "; QPTrim$(EmpData2FileRec.EmpNo); Tab(60); "Soc Sec No: "; QPTrim$(EmpData2FileRec.EmpSSN)
  Print #RptHandle, "    Last Name: "; QPTrim$(EmpData2FileRec.EmpLName); Tab(60); "First Name: "; QPTrim$(EmpData2FileRec.EmpFName)
  Print #RptHandle, "      Address: "; QPTrim$(EmpData2FileRec.EmpAddr1)
  Print #RptHandle, "      Address: "; QPTrim$(EmpData2FileRec.EMPADDR2)
  Print #RptHandle, "         City: "; QPTrim$(EmpData2FileRec.EmpCity); Tab(40); "State: "; QPTrim$(EmpData2FileRec.EmpState); Tab(60); "Zip: "; EmpData2FileRec.EmpZip
  Print #RptHandle, "    Birthdate: "; BDate$; Tab(40); "Gender: "; QPTrim$(EmpData2FileRec.EMPGENDR); Tab(60); "Race: "; QPTrim$(EmpData2FileRec.EMPRACE)
  Print #RptHandle, "   Ret Number: "; QPTrim$(EmpData2FileRec.EMPRETNO); Tab(40); "Ret Type: "; QPTrim$(EmpData2FileRec.EMPRETTP)
  Print #RptHandle, ""
  Print #RptHandle, " Direct Deposit Information"
  Print #RptHandle, "   BankDraft Code: "; QPTrim$(EmpData2FileRec.DRAFTCOD)
  Print #RptHandle, "     Bank Acct No: "; QPTrim$(EmpData2FileRec.EMPDDACC)
  Print #RptHandle, "         Prenoted: "; QPTrim$(EmpData2FileRec.PRENOTED)
  Print #RptHandle, "        Bank Name: "; QPTrim$(EmpData2FileRec.BankName)
  Print #RptHandle, "    Bank Location: "; QPTrim$(EmpData2FileRec.BANKLOC)
  Print #RptHandle, "  Bank Transit No: "; QPTrim$(EmpData2FileRec.TRANSIT)
  Print #RptHandle, ""
  Print #RptHandle, "  Job Description"
  Print #RptHandle, "        Title: "; QPTrim$(EmpData2FileRec.EMPJOB); Tab(47); "W/C Code: "; QPTrim$(EmpData2FileRec.EMPWCCLS)
  Print #RptHandle, "       Status: "; QPTrim$(EmpData2FileRec.EMPSTATS); Tab(36); "Benefit Pct: "; Using("##0.00", EmpData2FileRec.EMPBCODE); Tab(65); "Pay Type: "; QPTrim$(EmpData2FileRec.EMPPTYPE)
  Print #RptHandle, "    Frequency: "; QPTrim$(EmpData2FileRec.EMPPFREQ); Tab(36); "Rate: "; Using("##,##0.00", EmpData2FileRec.EMPPRATE); Tab(65); "O/T  Rate: "; Using("##,##0.00", EmpData2FileRec.EMPORATE)
  Print #RptHandle, "    Hire Date: "; HDate; Tab(36); "Next Review: "; RDate; Tab(65); "Term Date: "; TDate
  Print #RptHandle, "      Comment: "; QPTrim$(EmpData2FileRec.Comment)
  Print #RptHandle, ""
  Print #RptHandle, "Tax Withholding          Fixed"
  Print #RptHandle, "             Exempt  Amt/Pct  Figure   Status   # Allowances   Addit W/H Amt"
  Print #RptHandle, "   Federal: "; Tab(16); QPTrim$(EmpData2FileRec.EMPFEDX); Tab(25); QPTrim$(EmpData2FileRec.EMPFEDO2); Tab(32); Using("##0.00", EmpData2FileRec.EMPFEDO1); Tab(42); QPTrim$(EmpData2FileRec.EMPFEDS); Tab(54); EmpData2FileRec.EMPFEDA; Tab(69); Using("##0.00", EmpData2FileRec.EMPFEDAA)
  Print #RptHandle, "   State: "; Tab(16); QPTrim$(EmpData2FileRec.EMPSTAX); Tab(25); QPTrim$(EmpData2FileRec.EMPSTAO2); Tab(32); Using("##0.00", EmpData2FileRec.EMPSTAO1); Tab(42); QPTrim$(EmpData2FileRec.EMPSTAS); Tab(54); EmpData2FileRec.EMPSTAA; Tab(69); Using("##0.00", EmpData2FileRec.EMPSTAAA)
  Print #RptHandle, "   Social Security Exempt? "; QPTrim$(EmpData2FileRec.EMPSOCX); Tab(35); "Medicare Exempt? "; QPTrim$(EmpData2FileRec.EMPMEDX); Tab(63); "EIC Code: "; QPTrim$(EmpData2FileRec.EMPEIC)
  Print #RptHandle, ""
  Print #RptHandle, "   Misc Deductions   Amt/Pct   Figure  Inc O/T  Misc Deductions   Amt/Pct   Figure  Inc O/T  "
  Print #RptHandle, "   1. "; Desc$(1); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(1).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(1).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(1).DOTI);
  Print #RptHandle, "  2. "; Desc$(2); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(2).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(2).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(2).DOTI)
  Print #RptHandle, "   3. "; Desc$(3); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(3).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(3).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(3).DOTI);
  Print #RptHandle, "  4. "; Desc$(4); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(4).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(4).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(4).DOTI)
  Print #RptHandle, "   5. "; Desc$(5); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(5).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(5).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(5).DOTI);
  Print #RptHandle, "  6. "; Desc$(6); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(6).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(6).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(6).DOTI)
  Print #RptHandle, "   7. "; Desc$(7); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(7).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(7).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(7).DOTI);
  Print #RptHandle, "  8. "; Desc$(8); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(8).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(8).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(8).DOTI)
  Print #RptHandle, "   9. "; Desc$(9); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(9).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(9).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(9).DOTI);
  Print #RptHandle, " 10. "; Desc$(10); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(10).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(10).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(10).DOTI)
  Print #RptHandle, "  11. "; Desc$(11); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(11).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(11).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(11).DOTI);
  Print #RptHandle, " 12. "; Desc$(12); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(12).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(12).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(12).DOTI)
  Print #RptHandle, "  13. "; Desc$(13); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(13).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(13).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(13).DOTI);
  Print #RptHandle, " 14. "; Desc$(14); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(14).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(14).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(14).DOTI)
  Print #RptHandle, "  15. "; Desc$(15); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(15).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(15).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(15).DOTI);
  Print #RptHandle, " 16. "; Desc$(16); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(16).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(16).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(16).DOTI)
  Print #RptHandle, "  17. "; Desc$(17); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(17).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(17).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(17).DOTI);
  Print #RptHandle, " 18. "; Desc$(18); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(18).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(18).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(18).DOTI)
  Print #RptHandle, "  19. "; Desc$(19); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(19).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(19).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(19).DOTI);
  Print #RptHandle, " 20. "; Desc$(20); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(20).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(20).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(20).DOTI)
  Print #RptHandle, "  21. "; Desc$(21); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(21).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(21).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(21).DOTI);
  Print #RptHandle, " 22. "; Desc$(22); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(22).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(22).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(22).DOTI)
  Print #RptHandle, "  23. "; Desc$(23); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(23).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(23).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(23).DOTI);
  Print #RptHandle, " 24. "; Desc$(24); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(24).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(24).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(24).DOTI)
  Print #RptHandle, "  25. "; Desc$(25); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(25).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(25).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(25).DOTI);
  Print #RptHandle, " 26. "; Desc$(26); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(26).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(26).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(26).DOTI)
  Print #RptHandle, "  27. "; Desc$(27); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(27).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(27).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(27).DOTI);
  Print #RptHandle, " 28. "; Desc$(28); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(28).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(28).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(28).DOTI)
  Print #RptHandle, "  29. "; Desc$(29); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(29).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(29).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(29).DOTI);
  Print #RptHandle, " 30. "; Desc$(30); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(30).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(30).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(30).DOTI)
  Print #RptHandle, "  31. "; Desc$(31); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(31).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(31).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(31).DOTI);
  Print #RptHandle, " 32. "; Desc$(32); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(32).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(32).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(32).DOTI)
  Print #RptHandle, "  33. "; Desc$(33); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(33).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(33).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(33).DOTI);
  Print #RptHandle, " 34. "; Desc$(34); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(34).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(34).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(34).DOTI)
  Print #RptHandle, "  35. "; Desc$(35); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(35).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(35).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(35).DOTI);
  Print #RptHandle, " 36. "; Desc$(36); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(36).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(36).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(36).DOTI)
  Print #RptHandle, "  37. "; Desc$(37); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(37).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(37).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(37).DOTI);
  Print #RptHandle, " 38. "; Desc$(38); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(38).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(38).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(38).DOTI)
  Print #RptHandle, "  39. "; Desc$(39); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(39).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(39).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(39).DOTI);
  Print #RptHandle, " 40. "; Desc$(40); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(40).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(40).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(40).DOTI)
  Print #RptHandle, "  41. "; Desc$(41); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(41).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(41).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(41).DOTI);
  Print #RptHandle, " 42. "; Desc$(42); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(42).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(42).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(42).DOTI)
  Print #RptHandle, "  43. "; Desc$(43); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(43).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(43).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(43).DOTI);
  Print #RptHandle, " 44. "; Desc$(44); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(44).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(44).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(44).DOTI)
  Print #RptHandle, "  45. "; Desc$(45); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(45).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(45).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(45).DOTI);
  Print #RptHandle, " 46. "; Desc$(46); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(46).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(46).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(46).DOTI)
  Print #RptHandle, "  47. "; Desc$(47); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(47).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(47).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(47).DOTI);
  Print #RptHandle, " 48. "; Desc$(48); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(48).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(48).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(48).DOTI)
  Print #RptHandle, "  49. "; Desc$(49); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(49).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(49).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(49).DOTI);
  Print #RptHandle, " 50. "; Desc$(50); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(50).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(50).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(50).DOTI)
  Print #RptHandle, FF$
  Print #RptHandle, "--------------------------------------------------------------------------------------------"
  Print #RptHandle, ""
  Print #RptHandle, "                  C O N F I D E N T I A L   D A T A   F I L E  (C O N T)"
  Print #RptHandle, ""
  Print #RptHandle, "  Employee Information"
  Print #RptHandle, "       Number: "; QPTrim$(EmpData2FileRec.EmpNo); Tab(60); "Soc Sec No: "; QPTrim$(EmpData2FileRec.EmpSSN)
  Print #RptHandle, "    Last Name: "; QPTrim$(EmpData2FileRec.EmpLName); Tab(60); "First Name: "; QPTrim$(EmpData2FileRec.EmpFName)
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, "   Default Earning Codes    Account Number     Earnings"
  Print #RptHandle, "   1. "; Desc(51); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT1); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT1)
  Print #RptHandle, "   2. "; Desc(52); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT2); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT2)
  Print #RptHandle, "   3. "; Desc(53); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT3); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT3)
  Print #RptHandle, ""
  Print #RptHandle, "   Wage Account Numbers        Default Distribution"
  Print #RptHandle, "   1. "; QPTrim$(EDistAcct(1)); Tab(38); Using$("##0.00", EDistAmt(1))
  Print #RptHandle, "   2. "; QPTrim$(EDistAcct(2)); Tab(38); Using$("##0.00", EDistAmt(2))
  Print #RptHandle, "   3. "; QPTrim$(EDistAcct(3)); Tab(38); Using$("##0.00", EDistAmt(3))
  Print #RptHandle, "   4. "; QPTrim$(EDistAcct(4)); Tab(38); Using$("##0.00", EDistAmt(4))
  Print #RptHandle, "   5. "; QPTrim$(EDistAcct(5)); Tab(38); Using$("##0.00", EDistAmt(5))
  Print #RptHandle, "   6. "; QPTrim$(EDistAcct(6)); Tab(38); Using$("##0.00", EDistAmt(6))
  Print #RptHandle, "   7. "; QPTrim$(EDistAcct(7)); Tab(38); Using$("##0.00", EDistAmt(7))
  Print #RptHandle, "   8. "; QPTrim$(EDistAcct(8)); Tab(38); Using$("##0.00", EDistAmt(8))
  Print #RptHandle, ""
  Print #RptHandle, " Benefit Schedule          Earned        Used      Balance"
  Print #RptHandle, "     Vacation"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPVACE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPVUSED); Tab(54); Using$("##0.00", EmpData2FileRec.EMPVBAL)
  Print #RptHandle, "     Sick Leave"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPSLE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPSLUSE); Tab(54); Using$("##0.00", EmpData2FileRec.EMPSLBAL)
  Print #RptHandle, "     Comp Time"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPCTE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPCTUSE); Tab(54); Using$("##0.00", EmpData2FileRec.EMPCTBAL)
  Print #RptHandle, "     Personal"; Tab(29); Using$("##0.00", EmpData2FileRec.PERERN); Tab(42); Using$("##0.00", EmpData2FileRec.PerUsed); Tab(54); Using$("##0.00", EmpData2FileRec.PERBAL)
  Print #RptHandle, "     Holiday"; Tab(29); Using$("##0.00", EmpData2FileRec.HOLERN); Tab(42); Using$("##0.00", EmpData2FileRec.HolUsed); Tab(54); Using$("##0.00", EmpData2FileRec.HOLBAL)
  Print #RptHandle, ""
  Print #RptHandle, "     Leave Table"; Tab(20); EmpData2FileRec.LeaveTbl; Tab(30); "Exclude ESC"; Tab(45); EmpData2FileRec.ExcludeESC
  Print #RptHandle, FF$
  
  Return
End Sub

Private Sub PrintGraphics()
  Dim RptName As String, EDistAmt(1 To 8) As Double
  Dim Emp2RecLen As Integer, ECnt As Integer
  Dim DataFileSize As Long, cnt As Integer
  Dim DataNumOfRecs As Long, EDistAcct(1 To 8) As String
  Dim RptHandle As Integer
  Dim RptTitle As String
  Dim x As Integer
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim TDate$, HDate$, BDate$
  Dim RDate$, Nextx As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  ReDim Emp2Data(1) As EmpData2Type
  ReDim Desc(1 To 53) As String * 8
  ReDim DedCodeRec(1 To 50) As DedCodeRecType
  Dim dlm$
  Dim UHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim ThisCnt As Integer
  Dim EmpZip$, ZipLen As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  InFileNames(4) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  
  RptName$ = "PRRPTS\EMPDATAG.RPT"

  OpenDedCodeFile DedCodeFileHandle
  For cnt = 1 To 50
    Get DedCodeFileHandle, cnt, DedCodeRec(cnt)
    Desc$(cnt) = QPTrim$(DedCodeRec(cnt).DCDESC1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close DedCodeFileHandle
  OpenErnCodeFile ErnCodeFileHandle
  For cnt = 51 To 53
    Get ErnCodeFileHandle, cnt - 50, ErnCodeRec
    Desc$(cnt) = QPTrim$(ErnCodeRec.ERNCODE1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close ErnCodeFileHandle
 
  RptTitle$ = "Employee Information Listing"

  
  RptHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RptHandle
  OpenEmpData2File EmpData2FileHandle
  Get EmpData2FileHandle, thisRecordNum, EmpData2FileRec
  For ECnt = 1 To 8
     EDistAcct(ECnt) = EmpData2FileRec.EDist(ECnt).DAcct
     EDistAmt(ECnt) = EmpData2FileRec.EDist(ECnt).DAmt
  Next ECnt
  GoSub PrintEmpData
  Close EmpData2FileHandle
  Close RptHandle
  
  MainLog ("Employee Data File processed.")
  arEmpDataRpt.Show
'  frmLoadingRpt.Show

Exit Sub

PrintEmpData:
ThisCnt = ThisCnt + 1
EmpZip = QPTrim$(EmpData2FileRec.EmpZip) '06/08/04
ZipLen = Len(EmpZip) '06/08/04
If ZipLen > 5 Then
  EmpZip = Mid(EmpZip, 1, 5) + "-" + Mid(EmpZip, 6, ZipLen) '06/08/04
  EmpData2FileRec.EmpZip = EmpZip '06/08/04
End If

BDate$ = MakeRegDate(EmpData2FileRec.EMPBDAY)
If BDate = "12/31/1979" Then BDate = "No record"
HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
If HDate = "12/31/1979" Then HDate = "No record"
RDate = MakeRegDate(EmpData2FileRec.EMPRDATE)
If RDate = "12/31/1979" Then RDate = "No record"
TDate = MakeRegDate(EmpData2FileRec.EMPTDATE)
If TDate = "12/31/1979" Then TDate = "No record"
'                            0                                         1                                         2
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpNo); dlm; QPTrim$(EmpData2FileRec.EmpLName) & ", " & QPTrim$(EmpData2FileRec.EmpFName); dlm; QPTrim$(EmpData2FileRec.EmpSSN); dlm;
'                            3                                         4                                         5
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpAddr1); dlm; ""; dlm; QPTrim$(EmpData2FileRec.EMPADDR2); dlm;
'                            6                                         7                                         8
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpCity); dlm; QPTrim$(EmpData2FileRec.EmpState); dlm; QPTrim$(EmpData2FileRec.EmpZip); dlm;
'                   9                        10                                       11                                      12
Print #RptHandle, BDate$; dlm; QPTrim$(EmpData2FileRec.EMPGENDR); dlm; QPTrim$(EmpData2FileRec.EMPRACE); dlm; QPTrim$(EmpData2FileRec.EMPRETNO); dlm;
'                               13                                     14                                      15                                      16
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPRETTP); dlm; QPTrim$(EmpData2FileRec.DRAFTCOD); dlm; QPTrim$(EmpData2FileRec.EMPDDACC); dlm; QPTrim$(EmpData2FileRec.PRENOTED); dlm;
'                               17                                     18                                      19                                   20
Print #RptHandle, QPTrim$(EmpData2FileRec.BankName); dlm; QPTrim$(EmpData2FileRec.BANKLOC); dlm; QPTrim$(EmpData2FileRec.TRANSIT); dlm; QPTrim$(EmpData2FileRec.EMPJOB); dlm;
'                               21                                     22                                              23                                           24
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPWCCLS); dlm; QPTrim$(EmpData2FileRec.EMPSTATS); dlm; Using("##0.00", EmpData2FileRec.EMPBCODE); dlm; QPTrim$(EmpData2FileRec.EMPPTYPE); dlm;
'                               25                                     26                                                    27                                   28
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPPFREQ); dlm; Using("##,##0.00", EmpData2FileRec.EMPPRATE); dlm; Using("##,##0.00", EmpData2FileRec.EMPORATE); dlm; HDate; dlm;
'                   29          30                       31                                     32                                        33                                          34
Print #RptHandle, RDate; dlm; TDate; dlm; QPTrim$(EmpData2FileRec.EMPFEDX); dlm; QPTrim$(EmpData2FileRec.EMPFEDO2); dlm; Using("##0.00", EmpData2FileRec.EMPFEDO1); dlm; QPTrim$(EmpData2FileRec.EMPFEDS); dlm;
'                           35                                36                                               37                                      38
Print #RptHandle, EmpData2FileRec.EMPFEDA; dlm; Using("##0.00", EmpData2FileRec.EMPFEDAA); dlm; QPTrim$(EmpData2FileRec.EMPSTAX); dlm; QPTrim$(EmpData2FileRec.EMPSTAO2); dlm;
'                           39                                                40                                   41                                     42
Print #RptHandle, Using("##0.00", EmpData2FileRec.EMPSTAO1); dlm; QPTrim$(EmpData2FileRec.EMPSTAS); dlm; EmpData2FileRec.EMPSTAA; dlm; Using("##0.00", EmpData2FileRec.EMPSTAAA); dlm;
'                           43                                        44                                           45
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPSOCX); dlm; QPTrim$(EmpData2FileRec.EMPMEDX); dlm; QPTrim$(EmpData2FileRec.EMPEIC); dlm;
'46 - 245
For x = 1 To 50
  If QPTrim$(EmpData2FileRec.EmpDed(x).DPct) = "PERCENT" Then EmpData2FileRec.EmpDed(x).DPct = "PERCNT"
  Print #RptHandle, Desc$(x); dlm; EmpData2FileRec.EmpDed(x).DPct; dlm; Using$("###0.00", EmpData2FileRec.EmpDed(x).DAmt); dlm; QPTrim$(EmpData2FileRec.EmpDed(x).DOTI); dlm;
Next x

'                   246                           247                             248
Print #RptHandle, Desc(51); dlm; QPTrim$(EmpData2FileRec.EMPEACT1); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT1); dlm;
'                   249                           250                             251
Print #RptHandle, Desc(52); dlm; QPTrim$(EmpData2FileRec.EMPEACT2); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT2); dlm;
'                   252                           253                             254
Print #RptHandle, Desc(53); dlm; QPTrim$(EmpData2FileRec.EMPEACT3); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT3); dlm;
' 255 - 270
For x = 1 To 8
  Print #RptHandle, QPTrim$(EDistAcct(x)); dlm; Using$("##0.00", EDistAmt(x)); dlm;
Next x
'                                      271                                       272                                                  273
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPVACE); dlm; Using$("##0.00", EmpData2FileRec.EMPVUSED); dlm; Using$("##0.00", EmpData2FileRec.EMPVBAL); dlm;
'                                      274                                       275                                                  276
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPSLE); dlm; Using$("##0.00", EmpData2FileRec.EMPSLUSE); dlm; Using$("##0.00", EmpData2FileRec.EMPSLBAL); dlm;
'                                      277                                       278                                                  279
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPCTE); dlm; Using$("##0.00", EmpData2FileRec.EMPCTUSE); dlm; Using$("##0.00", EmpData2FileRec.EMPCTBAL); dlm;
'                                      280                                       281                                                  282
Print #RptHandle, Using$("##0.00", EmpData2FileRec.PERERN); dlm; Using$("##0.00", EmpData2FileRec.PerUsed); dlm; Using$("##0.00", EmpData2FileRec.PERBAL); dlm;
'                                      283                                       284                                                  285
Print #RptHandle, Using$("##0.00", EmpData2FileRec.HOLERN); dlm; Using$("##0.00", EmpData2FileRec.HolUsed); dlm; Using$("##0.00", EmpData2FileRec.HOLBAL); dlm;
'                          286                              287
Print #RptHandle, EmpData2FileRec.LeaveTbl; dlm; EmpData2FileRec.ExcludeESC; dlm; UnitRec.UFEMPR; dlm; EmpData2FileRec.YN401K; dlm; QPTrim$(EmpData2FileRec.Comment)

Return

ErrorHandler:
  Close
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Private Sub cmdSave_Click()
 If Val(Mid(txtNumber.Text, 1, 1)) = 0 Then
   MsgBox "Please enter another employee number that does not begin with a zero."
   vaTabPro1.ActiveTab = 0
   txtNumber.SetFocus
   Exit Sub
 End If
 
 InsertDashes2GLNum
 If CheckEmpNum = False Then
   Exit Sub
 End If
 
 If Check4EqualDist = True Then Exit Sub
 CheckForValidWHNum
  If BadGLNum = True Then
    BadGLNum = False
    Exit Sub
  End If
  Call SaveEmpInfo(newEmpFlag, thisRecordNum, Me)
  MainLog ("Employee data was saved.")
End Sub

Private Sub retNumCombo_LostFocus()
 If QPTrim$(fptxtRetNum.Text) = "" Then
   MsgBox "Please make an entry in the Ret Type field only if the Ret Number field is filled"
   fptxtRetNum.SetFocus
 End If

End Sub

Private Sub LoadEMFile()
'   On Error Resume
'
'  NOTE: a new employee is treated differently than an existing
'  employee in that an existing employee's records are loaded
'  on the forms
'  Entering this form from the Employee Maintenance Menu
'  form causes all three employee maintenance forms to load
'  but both 2 and 3 are hidden

'NOTE ...retooled on 8/13/02 to reduce redundancies
   Dim Today As String * 10
   Dim EmpData2FileHandle As Integer
   Dim EmpData2FileRec As EmpData2Type
   Dim EmpRecLen As Long
   Dim DedCodeFileRec As DedCodeRecType
   Dim DedCodeFileHandle As Integer
   Dim x As Integer
   Dim ThisState$
   Dim UnitHandle As Integer, UnitRec As UnitFileRecType
   Dim ErnCodeFileHandle As Integer
   Dim ErnCodeRec As ErnCodeRecType
   Dim LeaveHandle As Integer
   Dim LeaveRec As LeaveRecType
   Dim LeaveNum As Integer
   Dim RetFileHandle As Integer
   Dim RetRec As RetireRecType
   Dim RetNum As Integer
   Dim Nextx As Integer
   Dim SysRec As RegDSysFileRecType
   Dim SysHandle As Integer
   Dim DHandle As Integer
   Dim DraftRec As DraftInfoFileName
   Dim FileSize As Integer
   
   OpenPRDraftFile DHandle
   FileSize = LOF(DHandle) / Len(DraftRec)
   Close DHandle
   
   If FileSize = 0 Then
     LabelDD.Visible = True
   Else
     LabelDD.Visible = False
   End If
   OpenSysFile SysHandle
   Get SysHandle, 1, SysRec
   Close SysHandle
   
   chkTemp.Enabled = False
   chkRet.Enabled = False
   If QPTrim$(CurrCitiPath) = "" Then cmdList.Visible = False '8/12/04
   fptxtMainDept.Visible = False 'until further notice
   Label71.Visible = False 'until further notice
   
   'these next 6 lines hide the last two lines of
   'alternate earnings because we are planning to
   'add more two more lines to this screen
   fptxtD(4).Visible = False
   fptxtD(5).Visible = False
   fptxtAN(4).Visible = False
   fptxtAN(5).Visible = False
   fptxtE(4).Visible = False
   fptxtE(5).Visible = False

   RetireFlag = True 'assume this employee has a
   'retirement plan

   OpenLeaveFileName LeaveHandle
   LeaveNum = LOF(LeaveHandle) \ Len(LeaveRec)
   Close LeaveHandle

   OpenRetFile RetFileHandle
   RetNum = LOF(RetFileHandle) \ Len(RetRec)
   If RetNum = 0 Then RetireFlag = False
   'get 'em and load 'em
   For x = 1 To RetNum
   Get RetFileHandle, x, RetRec
     If Len(QPTrim$(RetRec.TYPEDES1)) > 0 Then
       fpcomboRetType.AddItem QPTrim$(Mid(RetRec.TYPEDES1, 1, 24))
     End If
   Next x
   Close RetFileHandle

'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   If frmEmployeeMaintMenu.Selection = neoOn Then 'the employee
   'maintenance menu command button for loading this screen
   'either goes through add new emp or edit existing emp...
   'if it goes through add new emp then a routine set up
   'there tells this code if the screen is set up for a new employee
   'or not
      newEmpFlag = True
      cmdMessage.Visible = False
      fpMaskBDay.Text = ""
      fpMaskHire.Text = Today
      fpMaskNext.Text = ""
      fpMaskTerm.Text = ""
   Else
'  RecNum is a global variable that is set in the Employee LookUp form
'  and is used extensively...a critical piece of this program
     thisRecordNum = RecNum
     newEmpFlag = False
     If CustHasMsg(RecNum) Then
       MsgAlertTimer.Enabled = True
     Else
       MsgAlertTimer.Enabled = False
       cmdMessage.ForeColor = &H80000012
     End If
   End If
   
    
   fpcomboGender.AddItem "Male"
   fpcomboGender.AddItem "Female"
   fpcomboBankdraft.InsertRow = "C" & Chr(9) & "Checking"
   fpcomboBankdraft.InsertRow = "S" & Chr(9) & "Savings"
   fpcomboPrenoted.AddItem "Y"
   fpcomboPrenoted.AddItem "N"
   fpcomboStatus.AddItem "Seasonal"
   fpcomboStatus.AddItem "Temporary"
   fpcomboStatus.AddItem "Full-Time"
   fpcomboStatus.AddItem "Part-Time"
   fpcomboPayType.AddItem "Salaried"
   fpcomboPayType.AddItem "Hourly"
   fpcomboFreq.AddItem "Weekly"
   fpcomboFreq.AddItem "Bi-Weekly"
   fpcomboFreq.AddItem "Semi-Monthly"
   fpcomboFreq.AddItem "Monthly"
   fpcomboFreq.AddItem "Quarterly"
   fpcomboFreq.AddItem "Semi-Annually"
   fpcomboFreq.AddItem "Annually"
   fpcombo401K.AddItem "Y"
   fpcombo401K.AddItem "N"
   
   fptxtHeader.Text = "For " & txtFirstName.Text & " " & txtLastName.Text
   OpenUnitFile UnitHandle
   Get UnitHandle, 1, UnitRec
   Close UnitHandle
   ThisState$ = QPTrim$(UnitRec.UFSTATE)

   OpenDedCodeFile DedCodeFileHandle
   DedCnt = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
'   These fields in the miscellaneous spreadsheet will always
'   be loaded regardless if this is a new employee or an
'   existing one

   vaSpreadMisc.MaxRows = DedCnt 'limit spreadsheet to
   'only list rows that have data

   For x = 1 To DedCnt
     Get DedCodeFileHandle, x, DedCodeFileRec
     vaSpreadMisc.Col = 1
     vaSpreadMisc.Row = x
     vaSpreadMisc.Text = QPTrim$(DedCodeFileRec.DCDESC1)
   Next x
   Close DedCodeFileHandle

   fpcomboFedX.AddItem "Y"
   fpcomboFedX.AddItem "N"
   fpcomboFedAmtPct.AddItem "A"
   fpcomboFedAmtPct.AddItem "P"
   fpcomboFedStatus.InsertRow = "S" & Chr$(3) & " Single"
   fpcomboFedStatus.InsertRow = "M" & Chr$(3) & " Married"
   fpcomboStateX.AddItem "Y"
   fpcomboStateX.AddItem "N"
   fpcomboStateAmtPct.AddItem "A"
   fpcomboStateAmtPct.AddItem "P"
   Select Case ThisState$
   Case "GA":
'     fpcomboStateStatus.InsertRow = "S" & Chr$(3) & ("GASingle")
'     fpcomboStateStatus.InsertRow = "M" & Chr$(3) & ("GAMarried")
'     fpcomboStateStatus.InsertRow = "H" & Chr$(3) & ("GAHead of HouseHold")
     fpcomboStateStatus.InsertRow = "H" & Chr$(3) & ("GASingle")
     fpcomboStateStatus.InsertRow = "G" & Chr$(3) & ("GAMarried")
     fpcomboStateStatus.InsertRow = "F" & Chr$(3) & ("GAHead of HouseHold")
   Case "SC":
     fpcomboStateStatus.InsertRow = "S" & Chr$(3) & ("SC Tax Table")
   Case "OK":
     fpcomboStateStatus.InsertRow = "S" & Chr$(3) & ("Single")
     fpcomboStateStatus.InsertRow = "M" & Chr$(3) & ("Married, Head of Household")
     fpcomboStateStatus.InsertRow = "D" & Chr$(3) & ("Dual Income Married")
   Case "AR":
     fpcomboStateStatus.InsertRow = "S" & Chr$(3) & ("Single")
     fpcomboStateStatus.InsertRow = "M" & Chr$(3) & ("Married (1 Exempt'n)")
     fpcomboStateStatus.InsertRow = "H" & Chr$(3) & ("Married/Head Fam(2 Exempt'ns")
   Case Else:
     fpcomboStateStatus.InsertRow = "S" & Chr$(3) & ("Single")
     fpcomboStateStatus.InsertRow = "M" & Chr$(3) & ("Married")
     fpcomboStateStatus.InsertRow = "H" & Chr$(3) & ("Head of HouseHold")
   End Select

   'find out how many leave tables exist and load
   'the leave table combo box with that many options
   fpcomboLT.AddItem "0"
   For x = 1 To LeaveNum
     fpcomboLT.AddItem (x)
   Next x
   fpcomboLT.MaxDrop = LeaveNum + 1
   
   fpcomboESC.AddItem "Y"
   fpcomboESC.AddItem "N"
   fpcomboSocX.AddItem "Y"
   fpcomboSocX.AddItem "N"
   fpcomboMedX.AddItem "Y"
   fpcomboMedX.AddItem "N"
   fpcomboEIC.InsertRow = "0" & Chr$(9) & "No Certificate"
   fpcomboEIC.InsertRow = "1" & Chr$(9) & "Employee Only"
   fpcomboEIC.InsertRow = "2" & Chr$(9) & "Employee & Spouse"
   OpenErnCodeFile ErnCodeFileHandle
   ' fields in the default earnings code spreadsheet are always
   ' loaded when this form is loaded
   For x = 1 To 3
     Get ErnCodeFileHandle, x, ErnCodeRec
     fptxtD(x).Text = QPTrim$(ErnCodeRec.ERNCODE1)
   Next x
   Close ErnCodeFileHandle
'  if this is an existing employee then load all his/her data
   If newEmpFlag = False Then
     cmdPrint.Visible = True
     OpenEmpData2File EmpData2FileHandle
     Get EmpData2FileHandle, thisRecordNum, EmpData2FileRec
     Close EmpData2FileHandle
     txtNumber.Text = QPTrim$(EmpData2FileRec.EmpNo)
     CurEmpNum = txtNumber.Text
     tempEmpNum = Val(QPTrim$(EmpData2FileRec.EmpNo))
     fpMaskSoc.Text = QPTrim$(EmpData2FileRec.EmpSSN)
     txtLastName.Text = QPTrim$(EmpData2FileRec.EmpLName)
     txtFirstName.Text = QPTrim$(EmpData2FileRec.EmpFName)
     fptxtHeader.Text = "For " & txtFirstName.Text & " " & txtLastName.Text
     txtAddress1.Text = QPTrim$(EmpData2FileRec.EmpAddr1)
     txtAddress2.Text = QPTrim$(EmpData2FileRec.EMPADDR2)
     txtCity.Text = QPTrim$(EmpData2FileRec.EmpCity)
     txtState.Text = QPTrim$(EmpData2FileRec.EmpState)
     txtZip.Text = QPTrim$(EmpData2FileRec.EmpZip)
     fpcomboGender.Text = QPTrim$((EmpData2FileRec.EMPGENDR))
     fptxtRace.Text = QPTrim$(EmpData2FileRec.EMPRACE)
     If fpMaskBDay.IsNull Then
       fpMaskBDay.DateValue = "00-00-0000"
       GoTo NullBDay
     End If
       If EmpData2FileRec.EMPBDAY = 0 Then
       fpMaskBDay.Text = ""
     Else
       fpMaskBDay.Text = MakeRegDate(EmpData2FileRec.EMPBDAY)
     End If
     If CheckValDate(fpMaskBDay.Text) = False Then
       fpMaskBDay.Text = ""
     End If
NullBDay:
     fptxtRetNum.Text = QPTrim$(EmpData2FileRec.EMPRETNO)
     If Mid(fptxtRetNum.Text, 1, 1) = "R" Then
       chkRet.Enabled = True
       chkTemp.Enabled = True
       chkRet.Value = 1
     ElseIf Mid(fptxtRetNum.Text, 1, 1) = "T" Then
       chkTemp.Enabled = True
       chkRet.Enabled = True
       chkTemp.Value = 1
     End If
     fpcomboRetType.Text = QPTrim$(EmpData2FileRec.EMPRETTP)
     '*****added 11/11/02
     fptxtHomePhone.Text = QPTrim$(EmpData2FileRec.HomePhone)
     fptxtContactPhone.Text = QPTrim$(EmpData2FileRec.EmrgncyCntctPhnNum)
     fptxtContactName.Text = QPTrim$(EmpData2FileRec.EmrgncyCntctName)
     fptxtRelationship.Text = QPTrim$(EmpData2FileRec.EmrgncyCntctRelation)
     
     fpcomboBankdraft.Text = QPTrim$(EmpData2FileRec.DRAFTCOD)
     txtBankAcctNo.Text = QPTrim$(EmpData2FileRec.EMPDDACC)
     fpcomboPrenoted.Text = QPTrim$(EmpData2FileRec.PRENOTED)
     txtBankName.Text = QPTrim$(EmpData2FileRec.BankName)
     txtBankLocation.Text = QPTrim$(EmpData2FileRec.BANKLOC)
'      some of the old records were assigning 0 to an empty variable
     txtBankTransNo.Text = QPTrim$(EmpData2FileRec.TRANSIT)
     
     If txtBankTransNo.Text = "0" Then
       txtBankTransNo.Text = ""
     End If

     txtTitle.Text = QPTrim$(EmpData2FileRec.EMPJOB)
     fptxtWCCode.Text = QPTrim$(EmpData2FileRec.EMPWCCLS)
     fpcomboStatus.Text = QPTrim$(EmpData2FileRec.EMPSTATS)
     fptxtBenefitPct.Text = Using("##0.00", EmpData2FileRec.EMPBCODE) & "%"
     fpcomboPayType.Text = QPTrim$(EmpData2FileRec.EMPPTYPE)
     fpcomboFreq.Text = QPTrim$(EmpData2FileRec.EMPPFREQ)
     fptxtRate.Text = EmpData2FileRec.EMPPRATE
     FormatCurrency (fptxtRate.Text)
     fptxtOTRate.Text = EmpData2FileRec.EMPORATE
     FormatCurrency (fptxtOTRate.Text)
     If fpMaskHire.IsNull Then
       fpMaskHire.DateValue = "00-00-0000"
       GoTo BadHireData
     End If
     If EmpData2FileRec.EMPHDATE = 0 Or CheckValDate(fpMaskHire.Text) = False Then
       fpMaskHire.Text = ""
     Else
       fpMaskHire.Text = MakeRegDate(EmpData2FileRec.EMPHDATE)
     End If
BadHireData:
     If fpMaskNext.IsNull Then
       fpMaskNext.DateValue = "00-00-0000"
       GoTo NullNextDate
     End If
     If EmpData2FileRec.EMPRDATE = 0 Then
       fpMaskNext.Text = ""
     Else
       fpMaskNext.Text = MakeRegDate(EmpData2FileRec.EMPRDATE)
     End If
     If CheckValDate(fpMaskNext.Text) = False Then
       fpMaskNext.Text = ""
     End If
NullNextDate:

     If fpMaskTerm.IsNull Then
       fpMaskTerm.DateValue = "00-00-0000"
       GoTo NullTermDate
     End If
     If EmpData2FileRec.EMPTDATE = 0 Then
       fpMaskTerm.Text = ""
     Else
       fpMaskTerm.Text = MakeRegDate(EmpData2FileRec.EMPTDATE)
     End If
     If CheckValDate(fpMaskTerm.Text) = False Then
       fpMaskTerm.Text = ""
     End If

NullTermDate:
     fptxtComment.Text = QPTrim$(EmpData2FileRec.Comment) 'added 9/1/04

     If QPTrim$(EmpData2FileRec.EMPPTYPE) = "Salaried" Then
       lblHrSal.Caption = "Enter Percentage"
     ElseIf QPTrim$(EmpData2FileRec.EMPPTYPE) = "Hourly" Then
       lblHrSal.Caption = "Hours Per Pay Period"
     End If
     
     fpcomboFedX.Text = QPTrim$(UCase$(EmpData2FileRec.EMPFEDX))
     fpcomboFedAmtPct.Text = QPTrim$(EmpData2FileRec.EMPFEDO2)
     fptxtFedFig.Text = Using$("##,##0.00", EmpData2FileRec.EMPFEDO1)
   
     If QPTrim(EmpData2FileRec.EMPFEDS) = "H" Then
       MsgBox "Head of Household is no longer a federal status type"
       EmpData2FileRec.EMPFEDS = ""
     End If
   
     fpcomboFedStatus.Text = EmpData2FileRec.EMPFEDS
     fptxtAllowNumFed.Text = EmpData2FileRec.EMPFEDA       'num of allowance
     fptxtAddWHFed.Text = EmpData2FileRec.EMPFEDAA
     fpcomboStateX.Text = EmpData2FileRec.EMPSTAX
     fpcomboStateAmtPct.Text = QPTrim$(EmpData2FileRec.EMPSTAO2)
     fptxtStateFig.Text = Using$("##,##0.00", EmpData2FileRec.EMPSTAO1)
     fpcomboStateStatus.Text = QPTrim$(EmpData2FileRec.EMPSTAS)
     fptxtAllowNumState.Text = EmpData2FileRec.EMPSTAA
     fptxtAddWHState.Text = EmpData2FileRec.EMPSTAAA
     fpcomboSocX.Text = QPTrim$(EmpData2FileRec.EMPSOCX)
     fpcomboMedX.Text = QPTrim$(EmpData2FileRec.EMPMEDX)
     fpcomboEIC.Text = EmpData2FileRec.EMPEIC
   
     For x = 1 To DedCnt '8/5
       vaSpreadMisc.Col = 2
       vaSpreadMisc.Row = x
       vaSpreadMisc.Text = EmpData2FileRec.EmpDed(x).DPct
       vaSpreadMisc.Col = 3
       vaSpreadMisc.Row = x
       vaSpreadMisc.Text = EmpData2FileRec.EmpDed(x).DAmt
       vaSpreadMisc.Col = 4
       vaSpreadMisc.Row = x
       vaSpreadMisc.Text = EmpData2FileRec.EmpDed(x).DOTI
       If vaSpreadMisc.Text = "Y" Then vaSpreadMisc.Text = "YES"
       If vaSpreadMisc.Text = "N" Then vaSpreadMisc.Text = "NO"
     Next x

   'ALERT>>>Provisions should be made here for 5 entries

'*************************************************************
     MainLog (QPTrim$(EmpData2FileRec.EmpFName) + " " + QPTrim(EmpData2FileRec.EmpLName) + " edit screen accessed.")
     fptxtAN(1).Text = QPTrim$(EmpData2FileRec.EMPEACT1)
     fptxtE(1).Text = EmpData2FileRec.EMPEAMT1
     fptxtAN(2).Text = QPTrim$(EmpData2FileRec.EMPEACT2)
     fptxtE(2).Text = EmpData2FileRec.EMPEAMT2
     fptxtAN(3).Text = QPTrim$(EmpData2FileRec.EMPEACT3)
     fptxtE(3).Text = EmpData2FileRec.EMPEAMT3
'     fptxtAN(4).Text = QPTrim$(EmpData2FileRec.EMPEACT2) '****temporary until we add more options
'     fptxtE(4).Text = EmpData2FileRec.EMPEAMT2
'     fptxtAN(5).Text = QPTrim$(EmpData2FileRec.EMPEACT3) '****temporary until we add more options
'     fptxtE(5).Text = EmpData2FileRec.EMPEAMT3
     
     For x = 1 To 8
     'W = Wage, D = Distribution, A = Account, N = Number
                                 'D = Default, D = Distribution
       fptxtWDAN(x).Text = QPTrim$(EmpData2FileRec.EDist(x).DAcct)
       fptxtWDDD(x).Text = EmpData2FileRec.EDist(x).DAmt
     Next x
     
     fptxtEarned(1).Text = Format(EmpData2FileRec.EMPVACE, "##0.00")
     fptxtUsed(1).Text = Format(EmpData2FileRec.EMPVUSED, "##0.00")
     fptxtBal(1).Text = Format(EmpData2FileRec.EMPVBAL, "##0.00")
     fptxtEarned(2).Text = Format(EmpData2FileRec.EMPSLE, "##0.00")
     fptxtUsed(2).Text = Format(EmpData2FileRec.EMPSLUSE, "##0.00")
     fptxtBal(2).Text = Format(EmpData2FileRec.EMPSLBAL, "##0.00")
     fptxtEarned(3).Text = Format(EmpData2FileRec.EMPCTE, "##0.00")
     fptxtUsed(3).Text = Format(EmpData2FileRec.EMPCTUSE, "##0.00")
     fptxtBal(3).Text = Format(EmpData2FileRec.EMPCTBAL, "##0.00")
     fptxtEarned(4).Text = Format(EmpData2FileRec.PERERN, "##0.00")
     fptxtUsed(4).Text = Format(EmpData2FileRec.PerUsed, "##0.00")
     fptxtBal(4).Text = Format(EmpData2FileRec.PERBAL, "##0.00")
     fptxtEarned(5).Text = Format(EmpData2FileRec.HOLERN, "##0.00")
     fptxtUsed(5).Text = Format(EmpData2FileRec.HolUsed, "##0.00")
     fptxtBal(5).Text = Format(EmpData2FileRec.HOLBAL, "##0.00")
     fpcomboLT.Text = EmpData2FileRec.LeaveTbl
     fpcombo401K.Text = EmpData2FileRec.YN401K
     fpcomboESC.Text = QPTrim$(EmpData2FileRec.ExcludeESC)
   Else 'new employee
     cmdPrint.Visible = False
     fptxtContactPhone.Text = "(000)-000-0000"
     fptxtHomePhone.Text = "(000)-000-0000"
     lblHrSal.Caption = "Hours or %/Salary"
     fptxtComment.Text = "" 'added 9/1/04
     cmdHistory.Visible = False
     cmdYTD.Visible = False
     fptxtAN(1).Text = ""
     fptxtE(1).Text = ""
     fptxtAN(2).Text = ""
     fptxtE(2).Text = ""
     fptxtAN(3).Text = ""
     fptxtE(3).Text = ""
     fptxtAN(4).Text = ""
     fptxtE(4).Text = ""
     fptxtAN(5).Text = ""
     fptxtE(5).Text = ""
     For x = 1 To 8
       fptxtWDAN(x).Text = ""
       fptxtWDDD(x).Text = "0.00"
     Next x
     fptxtEarned(1).Text = "0.00"
     fptxtUsed(1).Text = "0.00"
     fptxtBal(1).Text = "0.00"
     fptxtEarned(2).Text = "0.00"
     fptxtUsed(2).Text = "0.00"
     fptxtBal(2).Text = "0.00"
     fptxtEarned(3).Text = "0.00"
     fptxtUsed(3).Text = "0.00"
     fptxtBal(3).Text = "0.00"
     fptxtEarned(4).Text = "0.00"
     fptxtUsed(4).Text = "0.00"
     fptxtBal(4).Text = "0.00"
     fptxtEarned(5).Text = "0.00"
     fptxtUsed(5).Text = "0.00"
     fptxtBal(5).Text = "0.00"
     fpcomboLT.Text = "0"
     fpcombo401K.Text = "N"
     fpcomboESC.Text = "N"
   End If
End Sub

Private Sub cmdYTD_Click()
  frmEmpYTDTot.Show vbModal, Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%Y"
      Call cmdYTD_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%M"
      Call cmdMessage_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%H"
      Call cmdHistory_Click
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdList_Click
      KeyCode = 0
    Case vbKeyPageUp:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 1
        fpcomboBankdraft.SetFocus
      ElseIf vaTabPro1.ActiveTab = 1 Then
        vaTabPro1.ActiveTab = 2
        txtTitle.SetFocus
      ElseIf vaTabPro1.ActiveTab = 2 Then
        vaTabPro1.ActiveTab = 3
        fpcomboFedX.SetFocus
      ElseIf vaTabPro1.ActiveTab = 3 Then
        vaTabPro1.ActiveTab = 4
        vaSpreadMisc.SetFocus
        vaSpreadMisc.Col = 2
        vaSpreadMisc.Row = 1
      ElseIf vaTabPro1.ActiveTab = 4 Then
        vaTabPro1.ActiveTab = 5
        fptxtAN(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 5 Then
        vaTabPro1.ActiveTab = 6
        fptxtWDAN(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 6 Then
        vaTabPro1.ActiveTab = 7
        fptxtEarned(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 7 Then
        vaTabPro1.ActiveTab = 0
        txtNumber.SetFocus
      End If
      KeyCode = 0
    Case vbKeyPageDown:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 7
        fptxtEarned(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 1 Then
        vaTabPro1.ActiveTab = 0
        txtNumber.SetFocus
      ElseIf vaTabPro1.ActiveTab = 2 Then
        vaTabPro1.ActiveTab = 1
        fpcomboBankdraft.SetFocus
      ElseIf vaTabPro1.ActiveTab = 3 Then
        vaTabPro1.ActiveTab = 2
        txtTitle.SetFocus
      ElseIf vaTabPro1.ActiveTab = 4 Then
        vaTabPro1.ActiveTab = 3
        fpcomboFedX.SetFocus
      ElseIf vaTabPro1.ActiveTab = 5 Then
        vaTabPro1.ActiveTab = 4
        vaSpreadMisc.SetFocus
        vaSpreadMisc.Col = 2
        vaSpreadMisc.Row = 1
      ElseIf vaTabPro1.ActiveTab = 6 Then
        vaTabPro1.ActiveTab = 5
        fptxtAN(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 7 Then
        vaTabPro1.ActiveTab = 6
        fptxtWDAN(1).SetFocus
      End If
      KeyCode = 0
  End Select
End Sub
Private Sub fpcomboRetType_Advance(ByVal Direction As Integer, AutoAdvance As Integer)
  If Direction = 1 Or Direction = 3 Then
    vaTabPro1.ActiveTab = 1
    fpcomboBankdraft.SetFocus
  End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub Form_Load()
  fpMaskBDay.AllowNull = True
  fpMaskHire.AllowNull = True
  fpMaskTerm.AllowNull = True
  fpMaskNext.AllowNull = True
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call FixSpread
  Call LoadEMFile
  If newEmpFlag = True Then
    Me.HelpContextID = hlpAddANewEmployee
  Else
    Me.HelpContextID = hlpEditViewEmployee
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub


Private Sub fpcombo401K_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcomboESC.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpcomboLT.SetFocus
    KeyCode = 0
  End If
  
End Sub

Private Sub fpcomboBankdraft_KeyDown(KeyCode As Integer, Shift As Integer)
  'if a user is tabbing thru any screen that has list boxes
  'or combo boxes the default settings allows the tab key to change the
  'value of data in these controls...this code is designed
  'to prevent the user from inadvertantly changing data
  'while tabbing
  If KeyCode = vbKeySpace Then
    fpcomboBankdraft.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboBankdraft.ListIndex = -1
  End If
  If fpcomboBankdraft.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtBankTransNo.SetFocus
        KeyCode = 0
      End If
    End If
  End If
  If KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 0
    txtNumber.SetFocus
    KeyCode = 0
  End If
End Sub

Private Sub fpcomboBankdraft_LostFocus()
  fpcomboBankdraft.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcomboEIC_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeySpace Then
    fpcomboEIC.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboEIC.ListIndex = -1
  End If
  If fpcomboEIC.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboFedX.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 4
    vaSpreadMisc.SetFocus
    vaSpreadMisc.Col = 2
    vaSpreadMisc.Row = 1
    vaSpreadMisc.SetActiveCell 2, 1
  End If
  
End Sub

Private Sub fpcomboEIC_LostFocus()
  fpcomboEIC.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboESC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboESC.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboESC.ListIndex = -1
  End If
  If fpcomboESC.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtEarned(1).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = 39 Then
      vaTabPro1.ActiveTab = 0
      txtNumber.SetFocus
      KeyCode = 0
    End If
  End If
  
End Sub

Private Sub fpcomboESC_LostFocus()
  If QPTrim$(fpcomboESC.Text) = "" Then fpcomboESC.Text = "N"
  fpcomboESC.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboFedAmtPct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFedAmtPct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFedAmtPct.ListIndex = -1
  End If
  If fpcomboFedAmtPct.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboFedAmtPct_LostFocus()
  fpcomboFedAmtPct.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboFedStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFedStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFedStatus.ListIndex = -1
  End If
  If fpcomboFedStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboFedStatus_LostFocus()
  fpcomboFedStatus.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboFedX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFedX.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFedX.ListIndex = -1
  End If
  If fpcomboFedX.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcomboEIC.SetFocus
        KeyCode = 0
      End If
    End If
  End If
  If KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 2
    txtTitle.SetFocus
    KeyCode = 0
  End If
  
End Sub

Private Sub fpcomboFedX_LostFocus()
  fpcomboFedX.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboFreq_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFreq.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFreq.ListIndex = -1
  End If
  If fpcomboFreq.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
 
End Sub

Private Sub fpcomboFreq_LostFocus()
  fpcomboFreq.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboGender_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboGender.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboGender.ListIndex = -1
  End If
  If fpcomboGender.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboGender_LostFocus()
  fpcomboGender.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboLT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboLT.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboLT.ListIndex = -1
  End If
  If fpcomboLT.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboLT_LostFocus()
  If QPTrim$(fpcomboLT.Text) = "" Then fpcomboLT.Text = 0
  fpcomboLT.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboMedX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboMedX.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboMedX.ListIndex = -1
  End If
  If fpcomboMedX.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboMedX_LostFocus()
  fpcomboMedX.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboPayType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPayType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPayType.ListIndex = -1
  End If
  If fpcomboPayType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboPayType_LostFocus()
  fpcomboPayType.Action = ActionClearSearchBuffer
  If QPTrim$(fpcomboPayType.Text) = "Salaried" Then
    lblHrSal.Caption = "Enter Percentage"
  ElseIf QPTrim$(fpcomboPayType.Text) = "Hourly" Then
    lblHrSal.Caption = "Enter Hours Per Period"
  End If
End Sub

Private Sub fpcomboPrenoted_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrenoted.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrenoted.ListIndex = -1
  End If
  If fpcomboPrenoted.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboPrenoted_LostFocus()
  fpcomboPrenoted.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboRetType_Change()
  If Len(QPTrim$(fpcomboRetType.Text)) > 0 Then
    If Len(QPTrim$(fptxtRetNum.Text)) = 0 Then
      MsgBox "Please make sure there is a retirement number entered before entering the retirement type."
      fpcomboRetType.Text = ""
      fptxtRetNum.SetFocus
      Exit Sub
    End If
  End If
  If Len(QPTrim$(fpcomboRetType.Text)) = 0 Then
    If Len(QPTrim$(fptxtRetNum.Text)) > 0 Then
      MsgBox "Please make sure there is a retirement type entered since a retirement number has been entered."
      fpcomboRetType.Text = ""
      fpcomboRetType.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub fpcomboRetType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboRetType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboRetType.ListIndex = -1
  End If
  If fpcomboRetType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtContactName.SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
  
End Sub

Private Sub fpcomboRetType_LostFocus()
  fpcomboRetType.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboSocX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboSocX.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboSocX.ListIndex = -1
  End If
  If fpcomboSocX.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboSocX_LostFocus()
  fpcomboSocX.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboStateAmtPct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboStateAmtPct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboStateAmtPct.ListIndex = -1
  End If
  If fpcomboStateAmtPct.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboStateAmtPct_LostFocus()
  fpcomboStateAmtPct.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboStateStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboStateStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboStateStatus.ListIndex = -1
  End If
  If fpcomboStateStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboStateStatus_LostFocus()
  fpcomboStateStatus.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboStateX_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboStateX.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboStateX.ListIndex = -1
  End If
  If fpcomboStateX.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboStateX_LostFocus()
  fpcomboStateX.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboStatus_Click()
' these statements cause the fptxtBenefitPct field to automatically
' load the 100% value if Fulltime is selected
  If fpcomboStatus.Text = "Full-Time" Then
     fptxtBenefitPct.Text = Using("##0.00", 100) & "%"
  ElseIf fpcomboStatus.Text = "Part-Time" Then
     fptxtBenefitPct.Text = Using("##0.00", 0) & "%"
  End If
End Sub

Private Sub fpcomboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboStatus.ListIndex = -1
  End If
  If fpcomboStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcomboStatus_LostFocus()
  fpcomboStatus.Action = ActionClearSearchBuffer

End Sub

Private Sub fpCurrency1_Change()

End Sub

Private Sub fpMaskBDay_LostFocus()
' even though the birth date is not a required field we still
' want only valid dates entered if a date is entered at all
  If fpMaskBDay = "" Then GoTo noEntry
  If CheckValDate(QPTrim$(fpMaskBDay.Text)) = False Then
     MsgBox "Please enter a valid Birthday date or delete current entry"
     fpMaskBDay.SetFocus
     GoTo noEntry
  End If

noEntry:
End Sub

Private Sub fpMaskHire_LostFocus()
' this is a required field so a valid date must be entered before
' moving on
  If fpMaskHire = "" Then GoTo noEntry
  If CheckValDate(QPTrim$(fpMaskHire.Text)) = False Then
     MsgBox "Please enter a valid Hire date"
     fpMaskHire.SetFocus
     GoTo noEntry
  End If

noEntry:
End Sub

Private Sub fpMaskNext_LostFocus()
' even though the next review date is not a required field we still
' want only valid dates entered if a date is entered at all
  If fpMaskNext = "" Then GoTo noEntry
  If CheckValDate(QPTrim$(fpMaskNext.Text)) = False Then
     MsgBox "Please enter a valid Next Review date or delete current entry"
     fpMaskNext.SetFocus
     GoTo noEntry
  End If

noEntry:

End Sub

Private Sub fpMaskSoc_GotFocus()
  SSN = QPTrim$(fpMaskSoc.Text)
End Sub

Private Sub fpMaskSoc_LostFocus()
  Dim DoWhatFlag As SaveChangeOptions1
'featured disabled on 5/12/04
'  If SSNCheck(fpMaskSoc.Text) = True Then
'    DoWhatFlag = PromptBadSSNNum(Me)
'    Select Case DoWhatFlag
'    Case BadSSNNUMOption.badssnClose
'      fpMaskSoc.Text = SSN
'      fpMaskSoc.SetFocus
'      Exit Sub
'    Case BadSSNNUMOption.badssnOverride
'      fpMaskBDay.SetFocus
'      Exit Sub
'    End Select
'  End If
End Sub

Private Sub fpMaskTerm_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyRight Then
'    vaTabPro1.ActiveTab = 3
'    fpcomboFedX.SetFocus
'    KeyCode = 0
'  ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
'    txtTitle.SetFocus
'    KeyCode = 0
'  ElseIf KeyCode = vbKeyUp Then
'    SendKeys "+{Tab}"
'    KeyCode = 0
'  End If
End Sub

Private Sub fpMaskTerm_LostFocus()
' even though the next term date is not a required field we still
' want only valid dates entered if a date is entered at all
  If fpMaskTerm = "" Then GoTo noEntry
  If CheckValDate(QPTrim$(fpMaskTerm.Text)) = False Then
    MsgBox "Please enter a valid Next Review date or delete current entry"
    fpMaskTerm.SetFocus
    GoTo noEntry
  End If

noEntry:

End Sub

Private Sub fptxtAddWHFed_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fptxtAddWHState_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fptxtAllowNumFed_LostFocus()
  If CheckFor2ManyDecimals(fptxtAllowNumFed.Text) = True Then
    MsgBox "Invalid number entered"
    fptxtAllowNumFed.SetFocus
    Exit Sub
  End If
  If Len(QPTrim$(fptxtAllowNumFed.Text)) = 0 Then
    fptxtAllowNumFed = "0"
  End If
End Sub

Private Sub fptxtAllowNumState_LostFocus()
  If CheckFor2ManyDecimals(fptxtAllowNumState.Text) = True Then
    MsgBox "Invalid number entered"
    fptxtAllowNumState.SetFocus
    Exit Sub
  End If
  If Len(QPTrim$(fptxtAllowNumState.Text)) = 0 Then
    fptxtAllowNumState = "0"
  End If

End Sub

Private Sub fptxtAN_BeforeDropDown(Index As Integer, Cancel As Boolean)

End Sub

Private Sub fptxtAN_DblClick(Index As Integer, Button As Integer)
  fptxtAN(Index) = Clipboard.GetText
  Clipboard.Clear
  
End Sub

Private Sub fptxtAN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fptxtAN(Index).ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fptxtAN(Index).Text = ""
  End If
  If Index = 1 Then
    If KeyCode = vbKeyLeft Then
      vaTabPro1.ActiveTab = 4
      vaSpreadMisc.SetFocus
      vaSpreadMisc.SetActiveCell 2, 1
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      fptxtE(5).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  ElseIf Index <> 1 Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fptxtAN_LostFocus(Index As Integer)
  If Len(QPTrim(fptxtAN(Index).Text)) = 0 Then
    Exit Sub
  Else
    InsertDashes2GLNum
  End If

End Sub

Private Sub fptxtBenefitPct_LostFocus()
  If CheckFor2ManyDecimals(fptxtBenefitPct.Text) = True Then
    MsgBox "Invalid number entered"
    fptxtBenefitPct.SetFocus
    Exit Sub
  End If
  fptxtBenefitPct.Text = Using("##0.00", Val(ReplaceString(fptxtBenefitPct.Text, "%", ""))) & "%"

End Sub

Private Sub fptxtComment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 3
    fpcomboFedX.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
    txtTitle.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtContactPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 1
    fpcomboBankdraft.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    txtNumber.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fptxtHomePhone.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub fptxtE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 3 Then
    If KeyCode = vbKeyRight Then
      vaTabPro1.ActiveTab = 6
      fptxtWDAN(1).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      fptxtAN(1).SetFocus
      KeyCode = 0
    End If
  ElseIf Index <> 3 Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If
  
End Sub

Private Sub fptxtEarned_Advance(Index As Integer, ByVal Direction As Integer, AutoAdvance As Integer)
'  If Direction = 1 Then Direction = 0
End Sub

Private Sub fptxtEarned_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If Index = 1 Then
   If KeyCode = vbKeyLeft Then
     vaTabPro1.ActiveTab = 6
     fptxtWDAN(1).SetFocus
     KeyCode = 0
   ElseIf KeyCode = vbKeyUp Then
     fpcomboESC.SetFocus
     KeyCode = 0
   ElseIf KeyCode = vbKeyDown Then
     SendKeys "{Tab}"
     KeyCode = 0
   End If
 ElseIf Index <> 1 Then
   If KeyCode = vbKeyUp Then
     SendKeys "+{Tab}"
     KeyCode = 0
   ElseIf KeyCode = vbKeyDown Then
     SendKeys "{Tab}"
     KeyCode = 0
   End If
 End If
    
End Sub

Private Sub fptxtEarned_LostFocus(Index As Integer)
  Dim EARN As Double
  Dim Used As Double
  Dim Bal As Double
  
  EARN = Val(fptxtEarned(Index).Text)
  Used = Val(fptxtUsed(Index).Text)
  Bal = OldRound(EARN - Used)
  fptxtBal(Index).Text = Format(Bal, "##0.00")
  fptxtEarned(Index).Text = Format(EARN, "##0.00")
  
  If CheckFor2ManyDecimals(fptxtEarned(Index).Text) = True Then
    MsgBox "Invalid number entered"
    fptxtEarned(Index).SetFocus
    Exit Sub
  End If
End Sub

Private Sub fptxtFedFig_LostFocus()
  fptxtFedFig.Text = Format(fptxtFedFig.Text, "##,##0.00")
  If CheckFor2ManyDecimals(fptxtFedFig.Text) = True Then
    MsgBox "Invalid number entered"
    fptxtFedFig.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtFedFig.Text) = "0.00" Then
    Exit Sub
  ElseIf Len(QPTrim$(fptxtFedFig.Text)) = 0 Then
    fptxtFedFig.Text = "0.00"
  End If
  
  If Val(QPTrim$(fptxtFedFig.Text)) > 0 Then
    If Len(QPTrim$(fpcomboFedAmtPct.Text)) = 0 Then
      MsgBox "Please make an entry in the Federal Amt/Pct field before making an entry in the Federal Figure field."
      fptxtFedFig.Text = "0.00"
      fpcomboFedAmtPct.SetFocus
      Exit Sub
    End If
  End If

End Sub

Private Sub fptxtRetNum_Change()
  If QPTrim(fptxtRetNum.Text) <> "" Then
    chkRet.Enabled = True
    chkTemp.Enabled = True
  Else
    chkRet.Enabled = False
    chkTemp.Enabled = False
  End If
'  If Mid(fptxtRetNum.Text, 1, 1) = "T" Then
'    MsgBox "Placing a 'T' in front of the retirement number will prevent this employee from participating in any state retirement program."
'  End If

End Sub

Private Sub fptxtRetNum_LostFocus()
  If RetireFlag = False And Len(QPTrim$(fptxtRetNum.Text)) > 0 Then
    MsgBox "There are no Retirement options available."
    fptxtRetNum.Text = ""
    Exit Sub
  End If
  
End Sub

Private Sub fptxtStateFig_LostFocus()
  fptxtStateFig.Text = Format(fptxtStateFig.Text, "##,##0.00")
  If CheckFor2ManyDecimals(fptxtStateFig.Text) = True Then
    MsgBox "Invalid number entered"
    fptxtStateFig.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtStateFig.Text) = "0.00" Then
    Exit Sub
  ElseIf Len(QPTrim$(fptxtStateFig.Text)) = 0 Then
    fptxtStateFig.Text = "0.00"
  End If
  
  If Val(QPTrim$(fptxtStateFig.Text)) > 0 Then
    If Len(QPTrim$(fpcomboStateAmtPct.Text)) = 0 Then
      MsgBox "Please make an entry in the State Amt/Pct field before making an entry in the Federal Figure field."
      fptxtStateFig.Text = "0.00"
      fpcomboStateAmtPct.SetFocus
      Exit Sub
    End If
  End If

End Sub

Private Sub fptxtUsed_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 5 Then
    If KeyCode = vbKeyDown Then
      fpcomboLT.SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
  
End Sub

Private Sub fptxtUsed_LostFocus(Index As Integer)
  Dim EARN As Double
  Dim Used As Double
  Dim Bal As Double
  
  EARN = Val(fptxtEarned(Index).Text)
  Used = Val(fptxtUsed(Index).Text)
  Bal = OldRound(EARN - Used)
  fptxtBal(Index).Text = Format(Bal, "##0.00")
  fptxtUsed(Index).Text = Format(Used, "##0.00")
  
  If CheckFor2ManyDecimals(fptxtUsed(Index).Text) = True Then
    MsgBox "Invalid number entered"
    fptxtUsed(Index).SetFocus
    Exit Sub
  End If
End Sub

Private Sub fptxtWDAN_BeforeDropDown(Index As Integer, Cancel As Boolean)

End Sub

Private Sub fptxtWDAN_DblClick(Index As Integer, Button As Integer)
  fptxtWDAN(Index) = Clipboard.GetText
  Clipboard.Clear

End Sub

Private Sub fptxtWDAN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeySpace Then
    fptxtWDAN(Index).ListDown = True
    KeyCode = 0
  End If
  If KeyCode = vbKeyDelete Then
    fptxtWDAN(Index).Text = ""
  End If
  If Index = 1 Then
    If KeyCode = vbKeyLeft Then
      vaTabPro1.ActiveTab = 5
      fptxtAN(1).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      fptxtWDDD(8).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  ElseIf Index <> 1 Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fptxtWDAN_LostFocus(Index As Integer)
  If Len(QPTrim(fptxtWDAN(Index).Text)) = 0 Then
    Exit Sub
  Else
    InsertDashes2GLNum
  End If
End Sub

Private Sub fptxtWDDD_Advance(Index As Integer, ByVal Direction As Integer, AutoAdvance As Integer)
  If Index = 8 And Direction = 1 Then
    vaTabPro1.ActiveTab = 7
    fptxtEarned(1).SetFocus
    Direction = 0
  End If

End Sub

Private Sub fptxtWDDD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 8 Then
    If KeyCode = vbKeyRight Then
      vaTabPro1.ActiveTab = 7
      fptxtEarned(1).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      fptxtWDAN(1).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  ElseIf Index <> 8 Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fptxtWDDD_LostFocus(Index As Integer)
  If CheckFor2ManyDecimals(fptxtWDDD(Index).Text) = True Then
    MsgBox "Invalid number entered"
    fptxtWDDD(Index).SetFocus
    Exit Sub
  End If
  If QPTrim$(fptxtWDDD(Index).Text) = "" Then fptxtWDDD(Index).Text = "0"
End Sub


Private Sub txtBankAcctNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If fpcomboBankdraft.Text = "" Then
     MsgBox "Bank Acct No should only be filled in if BankDraft Code is filled in."
     txtBankAcctNo.Text = ""
     fpcomboBankdraft.SetFocus
   End If
     
End Sub

Private Sub txtBankLocation_KeyDown(KeyCode As Integer, Shift As Integer)
   If fpcomboBankdraft.Text = "" Then
     MsgBox "Bank Location should only be filled in if BankDraft Code is filled in."
     txtBankLocation.Text = ""
     fpcomboBankdraft.SetFocus
   End If

End Sub

Private Sub txtBankName_KeyDown(KeyCode As Integer, Shift As Integer)
   If fpcomboBankdraft.Text = "" Then
     MsgBox "Bank Name should only be filled in if BankDraft Code is filled in."
     txtBankName.Text = ""
     fpcomboBankdraft.SetFocus
   End If

End Sub

Private Sub txtBankTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 2
    txtTitle.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fpcomboBankdraft.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If
  
  If fpcomboBankdraft.Text = "" And Len(QPTrim$(txtBankTransNo.Text)) > 0 Then
    MsgBox "Bank Transit No should only be filled in if BankDraft Code is filled in."
    txtBankTransNo.Text = ""
    fpcomboBankdraft.SetFocus
  End If

End Sub

Private Sub txtFirstName_LostFocus()
' these statements load the header text ...effective
' only if this is a new employee or an existing employee whose name
' is being edited
  fptxtHeader.Text = "For " & txtFirstName.Text & " " & txtLastName.Text

End Sub

Private Sub txtLastName_LostFocus()
' these statements load the header text on forms ...effective
' only if this is a new employee or an existing employee whose name
' is being edited
  fptxtHeader.Text = "For " & txtFirstName.Text & " " & txtLastName.Text

End Sub


Private Sub txtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 7
    fptxtEarned(1).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpcomboRetType.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Function CheckEmpNum() As Boolean
  Dim Emp2Rec As EmpData2Type
  Dim EmpCnt As Integer
  Dim EmpHandle As Integer
  Dim x As Integer
  
  CheckEmpNum = True
  OpenEmpData2File EmpHandle
  EmpCnt = LOF(EmpHandle) / Len(Emp2Rec)
  If newEmpFlag = False Then
    For x = 1 To EmpCnt
      Get EmpHandle, x, Emp2Rec
      If x <> RecNum Then
        If QPTrim$(txtNumber.Text) = QPTrim$(Emp2Rec.EmpNo) Then
          CheckEmpNum = False
          MsgBox "This employee number is already in use. Please select another."
          'reload the current emp num in case the user
          'forgot what that number was
          txtNumber.Text = CurEmpNum '8/19 added
          vaTabPro1.ActiveTab = 0
          txtNumber.SetFocus
          Close EmpHandle
          Exit Function
        End If
      End If
    Next x
  Else
    For x = 1 To EmpCnt
      Get EmpHandle, x, Emp2Rec
        If QPTrim$(txtNumber.Text) = QPTrim$(Emp2Rec.EmpNo) Then
          CheckEmpNum = False
          MsgBox "This employee number is already in use. Please select another."
          vaTabPro1.ActiveTab = 0
          txtNumber.SetFocus
          Close EmpHandle
          Exit Function
        End If
    Next x
  End If
      
  Close EmpHandle
  
End Function

Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 1
    fpcomboBankdraft.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpMaskTerm.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub txtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub vaSpreadMisc_Click(ByVal Col As Long, ByVal Row As Long)
  
  If (Col > 1) Then
    vaSpreadMisc.EditMode = True
  Else
    vaSpreadMisc.EditMode = False
  End If
End Sub

Private Sub vaSpreadMisc_EditChange(ByVal Col As Long, ByVal Row As Long)
  '8/19 changed the error trap for incorrect entries in
  'columns 2 and 4 of the spreadsheet
  '8/7 added "Replace Existing Text" option in spreadsheet, located
  'in the spreadsheet edit under General Spreadsheet Environment...this
  'setting allows the focus to be on the entire cell which makes
  'it easier for the user to edit cell data
  If Col = 2 Then
    vaSpreadMisc.Col = Col
    vaSpreadMisc.Row = Row
    If QPTrim$(vaSpreadMisc.Text) = "A" Then vaSpreadMisc.Text = "AMOUNT"
    If QPTrim$(vaSpreadMisc.Text) = "P" Then vaSpreadMisc.Text = "PERCENT"
    If QPTrim$(vaSpreadMisc.Text) <> "AMOUNT" _
    And QPTrim$(UCase$(vaSpreadMisc.Text)) <> "PERCENT" _
    And QPTrim$(vaSpreadMisc.Text) <> "" Then
      MsgBox "Entry is not valid"
      vaSpreadMisc.Text = ""
      Exit Sub
    End If
  End If
  If Col = 4 Then
    vaSpreadMisc.Col = Col
    vaSpreadMisc.Row = Row
    If QPTrim$(vaSpreadMisc.Text) = "Y" Then vaSpreadMisc.Text = "YES"
    If QPTrim$(vaSpreadMisc.Text) = "N" Then vaSpreadMisc.Text = "NO"
    If QPTrim$(vaSpreadMisc.Text) <> "YES" _
    And QPTrim$(vaSpreadMisc.Text) <> "NO" _
    And QPTrim$(vaSpreadMisc.Text) <> "" Then
      MsgBox "Entry is not valid"
      vaSpreadMisc.Text = ""
      Exit Sub
    End If
 End If

End Sub

Private Sub vaSpreadMisc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown Then
    vaTabPro1.ActiveTab = 3
    fpcomboFedX.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyPageUp Then
    vaTabPro1.ActiveTab = 5
    fptxtAN(1).SetFocus
    KeyCode = 0
  End If
  
End Sub

Sub MakeGLDescIdx()
   Dim JGLIdxRec(1) As JGLAcctIdxType
   Dim GLIdxNum$
   Dim GLDHandle As Integer
   Dim GLIdxRecLen As Integer
   Dim GLDescRecLen As Integer
   Dim TotalAccts As Integer
   Dim Nextx As Integer, x As Integer
   Dim GLIDATDesc$, y As Integer
   Dim GLDesc(1) As GLAcctRecType
   Dim GLIdxHandle As Integer
   
   If Exist(GetCitiDirFolder + "GLACCT.IDX") Then
     GLIdxNum$ = GetCitiDirFolder + "GLACCT.IDX"
   Else
     Exit Sub
   End If
   
   If Exist(GetCitiDirFolder + "GLACCT.DAT") Then
     GLIDATDesc$ = GetCitiDirFolder + "GLACCT.DAT"
   Else
     Exit Sub
   End If
   
   GLIdxRecLen = Len(JGLIdxRec(1))
   GLDescRecLen = Len(GLDesc(1))
   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
   
   If TotalAccts = 0 Then Exit Sub
   
   ReDim DescBuff(1 To TotalAccts)
   GLIdxHandle = FreeFile
   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   GLDHandle = FreeFile
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   Nextx = 1
'   For x = Nextx To 5
'     For Y = 1 To TotalAccts
'       If DescBuff(Y) <> 0 Then
'         Get GLDHandle, DescBuff(Y), GLDesc(1)
'         fptxtAN(x).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       End If
'     Next Y
'     Nextx = Nextx + 1
'   Next x
  For y = 1 To TotalAccts
    If DescBuff(y) <> 0 Then
    Get GLDHandle, DescBuff(y), GLDesc(1)
       fptxtAN(1).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtAN(2).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtAN(3).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtAN(4).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtAN(5).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
   End If
   Next y
'   Nextx = 1
'   For x = Nextx To 8
'     For Y = 1 To TotalAccts
'       If DescBuff(Y) <> 0 Then
'         Get GLDHandle, DescBuff(Y), GLDesc(1)
'         fptxtWDAN(x).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       End If
'     Next Y
'     Nextx = Nextx + 1
'   Next x
'  For Y = 1 To TotalAccts
'    If DescBuff(Y) <> 0 Then
'    Get GLDHandle, DescBuff(Y), GLDesc(1)
'       fptxtWDAN(1).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(2).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(3).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(4).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(5).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(6).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(7).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'       fptxtWDAN(8).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & QPTrim$(GLDesc(1).Num) & Chr(9) & ReplaceString(GLDesc(1).Num, "-", "")
'   End If
'   Next Y
   
   Close GLDHandle

End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 18
        coladj = 10
        vaTabPro1.TabHeight = 500
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 10
        vaSpreadMisc.FontSize = 18
        vaSpreadMisc.RowHeight(-1) = 22
        vaSpreadMisc.RowHeight(0) = 22
      Else
        COne = 13
        coladj = 4
        vaSpreadMisc.RowHeight(-1) = 18
        vaSpreadMisc.RowHeight(0) = 18
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        vaTabPro1.TabHeight = 450
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 10
        COne = 14
        coladj = 7
        vaSpreadMisc.FontSize = 14
        vaSpreadMisc.RowHeight(0) = 18.5
        vaSpreadMisc.RowHeight(-1) = 18.5
      Else
        COne = 6.65
        coladj = 1.8
        vaSpreadMisc.RowHeight(0) = 16
        vaSpreadMisc.RowHeight(-1) = 17
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 13.49
        coladj = 3.1
        vaSpreadMisc.RowHeight(0) = 17.5
        vaSpreadMisc.RowHeight(-1) = 17.5
        vaSpreadMisc.FontBold = True
        For cnt = 1 To 8
          vaTabPro1.Tab = cnt - 1
          vaTabPro1.TabHeight = 450
          vaTabPro1.FontName = "Tahoma"
          vaTabPro1.FontSize = 12
          vaTabPro1.Tab = 7
        Next cnt
      Else
        COne = 1.2
        coladj = 0.35
      End If
      Case 800
        COne = 0
        coladj = 0
        vaSpreadMisc.Font.Size = 12
        vaSpreadMisc.RowHeight(-1) = 14
        For cnt = 1 To 7
          vaTabPro1.Tab = cnt - 1
          vaTabPro1.TabHeight = 350
          vaTabPro1.FontName = "Tahoma"
          vaTabPro1.FontSize = 8
        Next cnt
        'could not get tab 7 to change the font size!
        vaTabPro1.Tab = 7
        vaTabPro1.TabHeight = 350
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 8
      Case Else
      vaTabPro1.ActiveTabBold = False
    End Select
SkipAdjust:
    vaSpreadMisc.ColWidth(1) = vaSpreadMisc.ColWidth(1) + COne
    vaSpreadMisc.ColWidth(2) = vaSpreadMisc.ColWidth(2) + coladj
    vaSpreadMisc.ColWidth(3) = vaSpreadMisc.ColWidth(3) + coladj
    vaSpreadMisc.ColWidth(4) = vaSpreadMisc.ColWidth(4) + coladj

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEditEmpData.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub CheckForValidWHNum()
   Dim JGLIdxRec(1) As JGLAcctIdxType
   Dim GLIdxNum$
   Dim GLDHandle As Integer
   Dim GLIdxRecLen As Integer
   Dim GLDescRecLen As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim GLIDATDesc$
   Dim GLDesc(1) As GLAcctRecType
   Dim GLIdxHandle As Integer
   Dim DoWhatFlag As BadGLNUM2Option
   Dim n As Integer
   Dim FundLength As Integer
   Dim AcctLength As Integer
   Dim DetLength As Integer
   Dim SysHandle As Integer
   Dim SysRec As RegDSysFileRecType
   Dim Nextx As Integer
   Dim y As Integer
   Dim thisNum$
   
   On Error GoTo ERRORSTUFF1
   
   OpenSysFile SysHandle
   Get SysHandle, 1, SysRec
   Close SysHandle
   
   If QPTrim$(SysRec.GLCheckYN) = "N" Then Exit Sub
   'GetAcctStruct reads in general ledger numbers and distributes
   'this number into three fields based on three different lengths
'   Call GetAcctStruct(QPTrim$(SysRec.CITIDIR), FundLength, AcctLength, DetLength)
   Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
   If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
     Exit Sub
   End If
   On Error GoTo ERRORSTUFF2
   
   BadGLNum = False
   ExitFlag = False
   ListFlag = False
   
   On Error GoTo ERRORSTUFF3
'   If Exist(QPTrim$(SysRec.CITIDIR) + "GLACCT.IDX") Then
'     GLIdxNum$ = QPTrim$(SysRec.CITIDIR) + "GLACCT.IDX"
'   ElseIf Exist(QPTrim$(SysRec.CITIDIR) + "\GLACCT.IDX") Then
'     GLIdxNum$ = QPTrim$(SysRec.CITIDIR) + "\GLACCT.IDX"
'   Else
'     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
'     Exit Sub
'   End If
'   On Error GoTo ERRORSTUFF4
'
'   If Exist(QPTrim$(SysRec.CITIDIR) + "GLACCT.DAT") Then
'     GLIDATDesc$ = QPTrim$(SysRec.CITIDIR) + "GLACCT.DAT"
'   ElseIf Exist(QPTrim$(SysRec.CITIDIR) + "\GLACCT.DAT") Then
'     GLIDATDesc$ = QPTrim$(SysRec.CITIDIR) + "\GLACCT.DAT"
'   Else
'     MsgBox "No G/L account number validation possible...GLACCT.DAT could not be found."
'     Exit Sub
'   End If
   If Exist(QPTrim$(CurrCitiPath) + "GLACCT.IDX") Then
     GLIdxNum$ = QPTrim$(CurrCitiPath) + "GLACCT.IDX"
   ElseIf Exist(QPTrim$(CurrCitiPath) + "\GLACCT.IDX") Then
     GLIdxNum$ = QPTrim$(CurrCitiPath) + "\GLACCT.IDX"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.IDX could not be found."
     Exit Sub
   End If
   On Error GoTo ERRORSTUFF4

   If Exist(QPTrim$(CurrCitiPath) + "GLACCT.DAT") Then
     GLIDATDesc$ = QPTrim$(CurrCitiPath) + "GLACCT.DAT"
   ElseIf Exist(QPTrim$(CurrCitiPath) + "\GLACCT.DAT") Then
     GLIDATDesc$ = QPTrim$(CurrCitiPath) + "\GLACCT.DAT"
   Else
     MsgBox "No G/L account number validation possible...GLACCT.DAT could not be found."
     Exit Sub
   End If
   
   On Error GoTo ERRORSTUFF5
   
   GLIdxRecLen = Len(JGLIdxRec(1))
   GLDescRecLen = Len(GLDesc(1))
   TotalAccts = FileSize(GLIDATDesc$) \ GLDescRecLen
   
   If TotalAccts = 0 Then Exit Sub
   ReDim DescBuff(1 To TotalAccts)
   GLIdxHandle = FreeFile
   
   On Error GoTo ERRORSTUFF6
   
   Open GLIdxNum$ For Random As GLIdxHandle Len = GLIdxRecLen
   For x = 1 To TotalAccts
     Get GLIdxHandle, x, JGLIdxRec(1)
     DescBuff(x) = JGLIdxRec(1).RecNo
   Next x
   Close GLIdxHandle
   GLDHandle = FreeFile
   Open GLIDATDesc$ For Random As GLDHandle Len = GLDescRecLen
   
   'go thru each number one at a time and compare against all gl nums
   On Error GoTo ERRORSTUFF7
   For y = 1 To 5
      thisNum$ = QPTrim$(ReplaceString(fptxtAN(y).Text, "-", ""))
      If thisNum$ = "" Then GoTo ZeroText
      For x = 1 To TotalAccts
      If DescBuff(x) = 0 Then GoTo DescBuffIsZero
        Get GLDHandle, DescBuff(x), GLDesc(1)
          If thisNum$ = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
            Exit For
          End If
DescBuffIsZero:
       If x = TotalAccts Then
         DoWhatFlag = PromptBadGLNumVer2(Me)
         Select Case DoWhatFlag
         Case BadGLNUM2Option.badgl2Return
           Close
           vaTabPro1.ActiveTab = 5
           vaTabPro1.SetFocus
           fptxtAN(y).SetFocus
           BadGLNum = True
           Exit Sub
         Case BadGLNUM2Option.badgl2Save
           Close
           Exit Sub
         Case Else:
            'Do nothing because we don't know about any options except
            'save, review or abandon...used as a placeholder for adding
            'other options at a later date
         End Select
         Close GLDHandle
         Exit Sub
       End If
    Next x
ZeroText:
   Next y
   On Error GoTo ERRORSTUFF8
   
   For y = 1 To 8
      thisNum$ = QPTrim$(ReplaceString(fptxtWDAN(y).Text, "-", ""))
      If thisNum$ = "" Then GoTo NoText
      For x = 1 To TotalAccts
      If DescBuff(x) = 0 Then GoTo DescBuffIsZero1
        Get GLDHandle, DescBuff(x), GLDesc(1)
          If thisNum$ = QPTrim$(ReplaceString(GLDesc(1).Num, "-", "")) Then
            Exit For
          End If
DescBuffIsZero1:
       If x = TotalAccts Then
         DoWhatFlag = PromptBadGLNumVer2(Me)
         Select Case DoWhatFlag
         Case BadGLNUM2Option.badgl2Return
           Close
           BadGLNum = True
           vaTabPro1.SetFocus
           vaTabPro1.ActiveTab = 6
           fptxtWDAN(y).SetFocus
           Exit Sub
         Case BadGLNUM2Option.badgl2Save
           Close
           Exit Sub
         Case Else:
            'Do nothing because we don't know about any options except
            'save, review or abandon...used as a placeholder for adding
            'other options at a later date
         End Select
         Close GLDHandle
         Exit Sub
       End If
    Next x
NoText:
  Next y
  
  Close GLDHandle
  
Exit Sub

ERRORSTUFF1:
   MsgBox "ERROR1"
   GoTo ERRORSTUFF
   
ERRORSTUFF2:
   MsgBox "ERROR2"
   GoTo ERRORSTUFF
   
ERRORSTUFF3:
   MsgBox "ERROR3"
   GoTo ERRORSTUFF

ERRORSTUFF4:
   MsgBox "ERROR4"
   GoTo ERRORSTUFF

ERRORSTUFF5:
   MsgBox "ERROR5"
   GoTo ERRORSTUFF

ERRORSTUFF6:
   MsgBox "ERROR6"
   GoTo ERRORSTUFF

ERRORSTUFF7:
   MsgBox "ERROR7"
   GoTo ERRORSTUFF

ERRORSTUFF8:
   MsgBox "ERROR8"
   GoTo ERRORSTUFF

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEditEmpDatat", "CheckForValidWHNum", Erl)
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

Private Sub InsertDashes2GLNum()
  Dim FundLength As Integer
  Dim AcctLength As Integer
  Dim DetLength As Integer
  Dim x As Integer
  Dim SysHandle As Integer
  Dim SysRec As RegDSysFileRecType
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec
  Close SysHandle
  
'  Call GetAcctStruct(QPTrim$(SysRec.CITIDIR), FundLength, AcctLength, DetLength)
  Call GetAcctStruct(CurrCitiPath, FundLength, AcctLength, DetLength)
  If FundLength = 0 And AcctLength = 0 And DetLength = 0 Then
    Exit Sub
  End If

  For x = 1 To 8
    If Len(QPTrim$(fptxtWDAN(x).Text)) > 0 Then
      fptxtWDAN(x).Text = AddDashesToGLNumber(fptxtWDAN(x).Text, FundLength, AcctLength, DetLength)
    End If
  Next x
  
  For x = 1 To 3
    If Len(QPTrim$(fptxtAN(x).Text)) > 0 Then
      fptxtAN(x).Text = AddDashesToGLNumber(fptxtAN(x).Text, FundLength, AcctLength, DetLength)
    End If
  Next x
End Sub

Private Function Check4EqualDist() As Boolean
  Dim x As Integer
  Dim Num2Chk As Double
  Dim Nextx As Integer
  
  Check4EqualDist = False
  
  Nextx = 2
  Num2Chk = Val(fptxtWDDD(Nextx).Text)
  Do
    For x = 1 To 8
      If x <> Nextx Then
        If Val(fptxtWDDD(x).Text) <> 0 Then
          If Num2Chk = Val(fptxtWDDD(x).Text) Then
            MsgBox "There are equal distribution amounts that may cause small rounding problems. Please modify these entries slightly (ex. change 40/40 to 39.99/40.01)."
  '          fptxtWDDD(Nextx).Text = Val(fptxtWDDD(Nextx).Text) - 0.01
  '          fptxtWDDD(x).Text = Val(fptxtWDDD(x).Text) + 0.01
            Check4EqualDist = True
            vaTabPro1.ActiveTab = 6
            fptxtWDDD(x).SetFocus
            Exit Do
          End If
        End If
      End If
    Next x
    Nextx = Nextx + 1
    If Nextx = 9 Then Exit Do
    Num2Chk = Val(fptxtWDDD(Nextx).Text)
  Loop


End Function

Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdMessage.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdMessage.ForeColor = &H80000012
      cmdMessage.FontSize = BtnFnt
    Case 2
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 0.7
    Case 3
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 1.4
    Case 4
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.1
    Case 5
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.8
    Case 6
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 3.5
    Case 7
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 4.2
    Case 8
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 4.9
    Case 9
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
'  DoEvents
End Sub

Private Sub vaTabPro1_TabActivate(TabToActivate As Integer)
  
  Select Case TabToActivate
    Case 0:
      If newEmpFlag = True Then
        Me.HelpContextID = hlpAddANewEmployee
      Else
        Me.HelpContextID = hlpEditViewEmployee
      End If
    Case 1:
      Me.HelpContextID = hlpDirectDeposit
    Case 2:
      Me.HelpContextID = hlpJobDescription
    Case 3:
      Me.HelpContextID = hlpTaxWithholding
    Case 4:
      Me.HelpContextID = hlpMiscDeduction
    Case 5:
      Me.HelpContextID = hlpAlternate
    Case 6:
      Me.HelpContextID = hlpWageDistribution
    Case 7:
      Me.HelpContextID = hlpBenefit
    Case Else:
      
  End Select
    
    
End Sub
