VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmManualTransEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Transaction Entry"
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
   Icon            =   "frmManualTransEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   5625
      Left            =   675
      TabIndex        =   42
      Top             =   2190
      Width           =   10470
      _Version        =   196609
      _ExtentX        =   18478
      _ExtentY        =   9927
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   400
      TabsPerRow      =   5
      TabCount        =   5
      ShowFocusRect   =   0   'False
      TabMaxWidth     =   300
      ActiveTabBold   =   0   'False
      GrayAreaColor   =   13684944
      OffsetFromClientTop=   -1  'True
      HighestPrecedence=   1
      ShowEarMark     =   0   'False
      EarMarkColorDark=   13684944
      PageEarMarkColorDark=   13684944
      DataFormat      =   ""
      AutoSizeChildren=   3
      BookCornerGuardWidth=   60
      BookCornerGuardLength=   360
      ThreeDAppearance=   2
      DrawFocusRect   =   1
      DataField       =   ""
      TabCaption      =   "frmManualTransEntry.frx":08CA
      PageEarMarkPictureNext=   "frmManualTransEntry.frx":0CD7
      PageEarMarkPicturePrev=   "frmManualTransEntry.frx":0CF3
      EarMarkPictureNext=   "frmManualTransEntry.frx":0D0F
      EarMarkPicturePrev=   "frmManualTransEntry.frx":0D2B
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   4710
         Left            =   -24735
         TabIndex        =   57
         Top             =   -20160
         Width           =   9690
         _Version        =   196609
         _ExtentX        =   17092
         _ExtentY        =   8308
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmManualTransEntry.frx":0D47
         Begin EditLib.fpCurrency fptxtTotal 
            Height          =   396
            Left            =   4608
            TabIndex        =   38
            Top             =   3456
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtRWHData 
            Height          =   396
            Left            =   6912
            TabIndex        =   37
            Top             =   2784
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtSSWHData 
            Height          =   396
            Left            =   6912
            TabIndex        =   35
            Top             =   2208
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtFTWHData 
            Height          =   396
            Left            =   6912
            TabIndex        =   33
            Top             =   1632
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtOTData 
            Height          =   396
            Left            =   6912
            TabIndex        =   31
            Top             =   1056
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtMWHData 
            Height          =   396
            Left            =   2736
            TabIndex        =   36
            Top             =   2784
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtSTWHData 
            Height          =   396
            Left            =   2736
            TabIndex        =   34
            Top             =   2208
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtREGData 
            Height          =   396
            Left            =   2736
            TabIndex        =   32
            Top             =   1632
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin EditLib.fpCurrency fptxtGPData 
            Height          =   396
            Left            =   2736
            TabIndex        =   30
            Top             =   1056
            Width           =   1644
            _Version        =   196608
            _ExtentX        =   2900
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
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Tax:"
            ForeColor       =   &H8000000E&
            Height          =   348
            Left            =   3408
            TabIndex        =   78
            Top             =   3552
            Width           =   1116
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            FillColor       =   &H0080FFFF&
            Height          =   4380
            Left            =   192
            Top             =   192
            Width           =   9372
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Retirement W/H:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   5040
            TabIndex        =   65
            Top             =   2880
            Width           =   1836
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicare W/H:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   1008
            TabIndex        =   64
            Top             =   2928
            Width           =   1644
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Social Security W/H:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   4656
            TabIndex        =   63
            Top             =   2304
            Width           =   2268
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "State Tax W/H:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   960
            TabIndex        =   62
            Top             =   2304
            Width           =   1788
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Federal Tax W/H:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   4944
            TabIndex        =   61
            Top             =   1728
            Width           =   1980
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "OT:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   6432
            TabIndex        =   60
            Top             =   1152
            Width           =   444
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "REG:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2064
            TabIndex        =   59
            Top             =   1728
            Width           =   540
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Pay:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   1488
            TabIndex        =   58
            Top             =   1152
            Width           =   1164
         End
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   4710
         Left            =   -24750
         TabIndex        =   50
         Top             =   -20160
         Width           =   9705
         _Version        =   196609
         _ExtentX        =   17119
         _ExtentY        =   8308
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmManualTransEntry.frx":0D63
         Begin LpLib.fpCombo fpGLNum 
            Height          =   405
            Index           =   1
            Left            =   1830
            TabIndex        =   12
            Top             =   1290
            Width           =   4845
            _Version        =   196608
            _ExtentX        =   8546
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
            Object.TabStop         =   -1  'True
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   3
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   2
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
            ColumnEdit      =   1
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
            EditAlignH      =   2
            EditAlignV      =   0
            ColDesigner     =   "frmManualTransEntry.frx":0D7F
         End
         Begin LpLib.fpCombo fpGLNum 
            Height          =   405
            Index           =   2
            Left            =   1830
            TabIndex        =   14
            Top             =   1875
            Width           =   4845
            _Version        =   196608
            _ExtentX        =   8546
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
            Object.TabStop         =   -1  'True
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   3
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   2
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
            ColumnEdit      =   1
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
            EditAlignH      =   2
            EditAlignV      =   0
            ColDesigner     =   "frmManualTransEntry.frx":108A
         End
         Begin LpLib.fpCombo fpGLNum 
            Height          =   405
            Index           =   3
            Left            =   1830
            TabIndex        =   16
            Top             =   2445
            Width           =   4845
            _Version        =   196608
            _ExtentX        =   8546
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
            Object.TabStop         =   -1  'True
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   3
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   2
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
            ColumnEdit      =   1
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
            EditAlignH      =   2
            EditAlignV      =   0
            ColDesigner     =   "frmManualTransEntry.frx":1395
         End
         Begin LpLib.fpCombo fpGLNum 
            Height          =   405
            Index           =   4
            Left            =   1830
            TabIndex        =   18
            Top             =   3030
            Width           =   4845
            _Version        =   196608
            _ExtentX        =   8546
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
            Object.TabStop         =   -1  'True
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Text            =   ""
            Columns         =   3
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   2
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
            ColumnEdit      =   1
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
            EditAlignH      =   2
            EditAlignV      =   0
            ColDesigner     =   "frmManualTransEntry.frx":16A0
         End
         Begin EditLib.fpCurrency fptxtGLAmt 
            Height          =   396
            Index           =   1
            Left            =   7056
            TabIndex        =   13
            Top             =   1284
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
         Begin EditLib.fpCurrency fptxtGLAmt 
            Height          =   396
            Index           =   2
            Left            =   7056
            TabIndex        =   15
            Top             =   1872
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
         Begin EditLib.fpCurrency fptxtGLAmt 
            Height          =   396
            Index           =   3
            Left            =   7056
            TabIndex        =   17
            Top             =   2436
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
         Begin EditLib.fpCurrency fptxtGLAmt 
            Height          =   396
            Index           =   4
            Left            =   7056
            TabIndex        =   19
            Top             =   3012
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
         Begin VB.Shape Shape4 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4380
            Left            =   192
            Top             =   192
            Width           =   9372
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "2)"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   1200
            TabIndex        =   56
            Top             =   1920
            Width           =   348
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "4)"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   1200
            TabIndex        =   55
            Top             =   3072
            Width           =   348
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "3)"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   1200
            TabIndex        =   54
            Top             =   2496
            Width           =   348
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "1)"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   1200
            TabIndex        =   53
            Top             =   1344
            Width           =   348
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "GL Account Number"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   2880
            TabIndex        =   52
            Top             =   804
            Width           =   2364
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   7104
            TabIndex        =   51
            Top             =   804
            Width           =   1164
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   4710
         Left            =   390
         TabIndex        =   43
         Top             =   615
         Width           =   9705
         _Version        =   196609
         _ExtentX        =   17119
         _ExtentY        =   8308
         _StockProps     =   70
         BackColor       =   9405029
         BorderAlignTextH=   2
         Caption         =   ""
         AutoSizeOffsetLeft=   1
         AutoSizeOffsetRight=   1
         AutoSizeOffsetTop=   1
         AutoSizeOffsetBottom=   1
         Picture         =   "frmManualTransEntry.frx":19AB
         Begin EditLib.fpText fptxtCmpHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   7
            Top             =   2210
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtHrsWkd 
            Height          =   390
            Left            =   5472
            TabIndex        =   4
            Top             =   624
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtSckHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   5
            Top             =   1152
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtVacHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   6
            Top             =   1680
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtPerHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   8
            Top             =   2736
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtHolHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   9
            Top             =   3264
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin EditLib.fpText fptxtOTHrs 
            Height          =   390
            Left            =   5472
            TabIndex        =   10
            Top             =   3792
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1958
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
            CharValidationText=   ". 1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   10
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
         Begin VB.Shape Shape3 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4380
            Left            =   192
            Top             =   192
            Width           =   9372
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Personal Hours:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   76
            Top             =   2832
            Width           =   2124
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "OT Hours Paid:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   49
            Top             =   3888
            Width           =   2124
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Holiday Hours:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   48
            Top             =   3360
            Width           =   2124
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comp Hours:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   47
            Top             =   2304
            Width           =   2124
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vacation Worked:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   46
            Top             =   1776
            Width           =   2124
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sick Hours:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   45
            Top             =   1248
            Width           =   2124
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hours Worked:"
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   2832
            TabIndex        =   44
            Top             =   720
            Width           =   2124
         End
      End
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   4710
         Left            =   -24750
         TabIndex        =   66
         Top             =   -20160
         Width           =   9705
         _Version        =   196609
         _ExtentX        =   17119
         _ExtentY        =   8308
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmManualTransEntry.frx":19C7
         Begin EditLib.fpCurrency fpcurrTotDeds 
            Height          =   396
            Left            =   5184
            TabIndex        =   22
            Top             =   3888
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
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
            ControlType     =   1
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
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3270
            Left            =   1770
            TabIndex        =   21
            Top             =   390
            Width           =   6300
            _Version        =   196613
            _ExtentX        =   11113
            _ExtentY        =   5768
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
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
            MaxCols         =   2
            MaxRows         =   50
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            ShadowColor     =   13684944
            SpreadDesigner  =   "frmManualTransEntry.frx":19E3
            ScrollBarTrack  =   3
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Deductions:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   2880
            TabIndex        =   77
            Top             =   3984
            Width           =   2028
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4380
            Left            =   192
            Top             =   192
            Width           =   9372
         End
      End
      Begin ImpproLib.vaImprint vaImprint5 
         Height          =   4710
         Left            =   -24750
         TabIndex        =   67
         Top             =   -20160
         Width           =   9705
         _Version        =   196609
         _ExtentX        =   17119
         _ExtentY        =   8308
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmManualTransEntry.frx":1F1E
         Begin EditLib.fpCurrency fptxtRG 
            Height          =   396
            Left            =   5328
            TabIndex        =   29
            Top             =   3696
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtMG 
            Height          =   396
            Left            =   5328
            TabIndex        =   28
            Top             =   3168
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtSocG 
            Height          =   396
            Left            =   5328
            TabIndex        =   27
            Top             =   2640
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtStaG 
            Height          =   396
            Left            =   5328
            TabIndex        =   26
            Top             =   2112
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtFG 
            Height          =   396
            Left            =   5328
            TabIndex        =   25
            Top             =   1584
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtNP 
            Height          =   396
            Left            =   5328
            TabIndex        =   24
            Top             =   1056
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin EditLib.fpCurrency fptxtEIC 
            Height          =   396
            Left            =   5328
            TabIndex        =   23
            Top             =   528
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
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
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Retirement Gross:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   2544
            TabIndex        =   74
            Top             =   3792
            Width           =   2076
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Medicare Gross:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   2544
            TabIndex        =   73
            Top             =   3264
            Width           =   1836
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Social Gross"
            ForeColor       =   &H8000000E&
            Height          =   348
            Left            =   2544
            TabIndex        =   72
            Top             =   2736
            Width           =   1644
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "State Gross:"
            ForeColor       =   &H8000000E&
            Height          =   396
            Left            =   2544
            TabIndex        =   71
            Top             =   2208
            Width           =   1404
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Federal Gross:"
            ForeColor       =   &H8000000E&
            Height          =   348
            Left            =   2544
            TabIndex        =   70
            Top             =   1680
            Width           =   2268
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Net Pay:"
            ForeColor       =   &H8000000E&
            Height          =   348
            Left            =   2544
            TabIndex        =   69
            Top             =   1152
            Width           =   972
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Earned Income Credit:"
            ForeColor       =   &H8000000E&
            Height          =   348
            Left            =   2544
            TabIndex        =   68
            Top             =   624
            Width           =   2508
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   4380
            Left            =   192
            Top             =   192
            Width           =   9372
         End
      End
   End
   Begin EditLib.fpText fptxtCheckNum 
      Height          =   348
      Left            =   7968
      TabIndex        =   3
      Top             =   1680
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
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
      MaxLength       =   15
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
   Begin EditLib.fpText fpEmpNameNum 
      Height          =   396
      Left            =   3264
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   768
      Width           =   4668
      _Version        =   196608
      _ExtentX        =   8234
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
   Begin EditLib.fpDateTime fpBegDate 
      Height          =   372
      Left            =   2016
      TabIndex        =   0
      Top             =   1680
      Width           =   1692
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
      DateMin         =   "19800101"
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
   Begin EditLib.fpDateTime fptxtEnd 
      Height          =   372
      Left            =   3984
      TabIndex        =   1
      Top             =   1680
      Width           =   1692
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
      DateMin         =   "19800101"
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
   Begin EditLib.fpDateTime fpCheckDate 
      Height          =   372
      Left            =   5952
      TabIndex        =   2
      Top             =   1680
      Width           =   1692
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "10-01-2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm-dd-yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCont 
      Height          =   570
      Left            =   5640
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   "Press to include this employee payroll data to the batch file."
      Top             =   7950
      Width           =   2295
      _Version        =   131072
      _ExtentX        =   4048
      _ExtentY        =   1005
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
      ButtonDesigner  =   "frmManualTransEntry.frx":1F3A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   570
      Left            =   8205
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Press to remove this employee payroll data from the batch file."
      Top             =   7950
      Width           =   2295
      _Version        =   131072
      _ExtentX        =   4048
      _ExtentY        =   1005
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
      ButtonDesigner  =   "frmManualTransEntry.frx":211A
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date:"
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   6192
      TabIndex        =   75
      Top             =   1392
      Width           =   1452
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number:"
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   7824
      TabIndex        =   41
      Top             =   1392
      Width           =   1692
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Period End:"
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   4272
      TabIndex        =   40
      Top             =   1392
      Width           =   1308
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Period Beg:"
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   2304
      TabIndex        =   39
      Top             =   1392
      Width           =   1356
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manual Transaction Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2688
      TabIndex        =   11
      Top             =   336
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1464
      Top             =   172
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1464
      Top             =   48
      Width           =   8652
   End
End
Attribute VB_Name = "frmManualTransEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim HourlyFlag As Boolean
  Dim FirstTimeThru As Boolean
  Dim ManTrans(1) As ManualTransRecType
  Dim OTActive As Boolean
  Dim DedRecCnt As Integer

Public Sub cmdCont_Click()
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim TransRec(1) As TransRecType
  Dim THandle As Integer
  Dim x As Integer
  Dim ThisDate As Integer
  Dim ChkDate As Integer
  
  If CheckValDate(fpBegDate.Text) = False Then
    MsgBox "Please enter a valid date."
    fpBegDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fpCheckDate.Text) = False Then
    MsgBox "Please enter a valid date."
    fpCheckDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fptxtEnd.Text) = False Then
    MsgBox "Please enter a valid date."
    fptxtEnd.SetFocus
    Exit Sub
  End If
  
  If Date2Num(fptxtEnd.Text) < Date2Num(fpBegDate.Text) Then
    MsgBox "The ending date must come after the beginning date."
    fptxtEnd.SetFocus
    Exit Sub
  End If
  
  If Date2Num(fpCheckDate.Text) < Date2Num(fpBegDate.Text) Then
    MsgBox "Please enter a check date on or after the beginning date."
    fpCheckDate.SetFocus
    Exit Sub
  End If
  
  ThisDate = Date2Num(Date)
  ChkDate = Date2Num(fpCheckDate.Text)
  If ChkDate - ThisDate >= Abs(60) Then
    If MsgBox("The check date is more than 60 days from today's date. If you wish to edit this date then press Yes.", vbYesNo) = vbYes Then
      Close
      vaTabPro1.SetFocus
      vaTabPro1.ActiveTab = 0
      fpCheckDate.SetFocus
      Exit Sub
    Else
      MainLog "User warned that the check date " + fpCheckDate.Text + " is over 60 days away from today's date " + CStr(Date) + " and elected to continue anyway."
    End If
  End If
  
  OpenTransWorkFile TRHandle
  
  Get TRHandle, RecNum, TransRec(1)
  ParseScrn2Manual TransRec(), ManTrans()
  ParseManual2Trans TransRec(), ManTrans()
  
  TransRec(1).TActive = -1
  Put TRHandle, RecNum, TransRec(1)
  Close TRHandle
  
  PPDRec.PACTIVE = 0  'moved from below on 10/3/06
  
  PPDRec.MACTIVE = -1 'moved from below on 10/3/06

  OpenPPDefaultFile PHandle 'moved from below on 10/3/06
  Put PHandle, 1, PPDRec
  Close PHandle
  
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)
  CreateEmpTransRecs RecNum
  
  TransRec(1).EmpPin = RecNum 'added 1/4/07
  Put TRHandle, RecNum, TransRec(1)
  Close TRHandle
  
'  PPDRec.PACTIVE = 0  '
'
'  PPDRec.MACTIVE = -1
'
'  OpenPPDefaultFile PHandle
'  Put PHandle, 1, PPDRec
'  Close PHandle
  MainLog ("Manual Transaction Entry Continue command used.")
  frmTransEntryEdit.Show
  DoEvents
  Unload frmManualTransEntry
End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub cmdDelete_Click()
  Dim DoWhatFlag As PRTRemove
  Dim THandle As Integer, TRec As TransRecType
  
  DoWhatFlag = PromptPRTRemove(Me)
  Select Case DoWhatFlag
  Case PRTRemove.prtrEscape
     MainLog ("Manual transaction delete activated...escape chosen.")
     Exit Sub
  Case PRTRemove.prtrDelete
     Call DeleteThisEmp
     MainLog ("Manual transaction delete activated...process delete chosen.")
  End Select
  
  frmTransEntryEdit.Show
  DoEvents
  Unload frmManualTransEntry
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdCont_Click
      KeyCode = 0
    Case vbKeyPageUp:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 1
        fpGLNum(1).SetFocus
      ElseIf vaTabPro1.ActiveTab = 1 Then
        vaTabPro1.ActiveTab = 2
        vaSpread1.Col = 1
        vaSpread1.Row = 1
        vaSpread1.SetFocus
        vaSpread1.SetActiveCell 1, 1
      ElseIf vaTabPro1.ActiveTab = 2 Then
        vaTabPro1.ActiveTab = 3
        fptxtEIC.SetFocus
      ElseIf vaTabPro1.ActiveTab = 3 Then
        vaTabPro1.ActiveTab = 4
        fptxtGPData.SetFocus
      ElseIf vaTabPro1.ActiveTab = 4 Then
        vaTabPro1.ActiveTab = 0
        fptxtHrsWkd.SetFocus
      End If
    KeyCode = 0
    Case vbKeyPageDown:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 4
        fptxtTotal.SetFocus
      ElseIf vaTabPro1.ActiveTab = 1 Then
        vaTabPro1.ActiveTab = 0
        fptxtOTHrs.SetFocus
      ElseIf vaTabPro1.ActiveTab = 2 Then
        vaTabPro1.ActiveTab = 1
        fptxtGLAmt(4).SetFocus
      ElseIf vaTabPro1.ActiveTab = 3 Then
        vaTabPro1.ActiveTab = 2
        vaSpread1.SetFocus
        vaSpread1.SetActiveCell 1, 1
      ElseIf vaTabPro1.ActiveTab = 4 Then
        vaTabPro1.ActiveTab = 3
        fptxtRG.SetFocus
      End If
      KeyCode = 0
  Case Else:
  
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call FixSpread
  LoadMTEFile
  MainLog ("Manual Transaction data entry screen accessed.")
  Me.HelpContextID = hlpEnterManual
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Sub LoadMTEFile()
  Dim Emp2Rec(1) As EmpData2Type
  Dim Emp2Handle As Integer
  Dim EmpRecNum As Long, x As Long
  Dim DedRec As DedCodeRecType
  Dim DedHandle As Integer
  Dim DedRecNum As Integer
  Dim Today As String * 10
  Dim ScrWidth As Long
  Dim Image$
  Dim ManRecLen As Integer
  Dim TransRec(1) As TransRecType
  Dim THandle As Integer
  
  If ScrWidth = 800 Then
    vaTabPro1.FontBold = False
    vaTabPro1.FontSize = 10
    vaTabPro1.ActiveTabBold = True
    vaTabPro1.ApplyTo = 2
  End If
  
  ManRecLen = Len(ManTrans(1))
  Image$ = "$###,##0.00"
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  
  OpenEmpData2File Emp2Handle
  Get Emp2Handle, RecNum, Emp2Rec(1)
  Close Emp2Handle
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)
'  TransRec(1).NetPay = TransRec(1).NetPay
  If TransRec(1).TActive = 0 Then
    FirstTimeThru = True
    OTActive = False
    OpenDedCodeFile DedHandle
    DedRecCnt = LOF(DedHandle) \ Len(DedRec)
    Call LoadScrnTActiveIsFalse
'    CreateEmpTransRecs RecNum
    Close TRHandle
    fpEmpNameNum.Text = QPTrim$(Emp2Rec(1).EmpNo) + "   " + QPTrim$(Emp2Rec(1).EmpFName) + " " + QPTrim$(Emp2Rec(1).EmpLName)
    
    vaSpread1.MaxRows = DedRecCnt 'limit spreadsheet to
   'only list rows that have data
    
    For x = 1 To DedRecCnt
       Get DedHandle, x, DedRec
       vaSpread1.Col = 1
       vaSpread1.Row = x
       vaSpread1.Text = DedRec.DCDESC1
    Next x
    Close DedHandle
  
    Call MakeGLDescIdx
    Exit Sub
  Else
    OTActive = True
'    CreateEmpTransRecs RecNum
  End If
  Close TRHandle
  
  fpEmpNameNum.Text = QPTrim$(Emp2Rec(1).EmpNo) + "   " + QPTrim$(Emp2Rec(1).EmpFName) + " " + QPTrim$(Emp2Rec(1).EmpLName)
  
  OpenDedCodeFile DedHandle
  DedRecCnt = LOF(DedHandle) \ Len(DedRec)
  
'  If OTActive = True Then
'    CreateEmpTransRecs RecNum
'    CalcPay TransRec(1), RecNum, True
    ParseTrans2Manual TransRec(), ManTrans()
'  End If
  
  ParseManual2Scrn TransRec(), ManTrans()
  
'  Close TRHandle
  
  vaSpread1.MaxRows = DedRecCnt
  For x = 1 To DedRecCnt
     Get DedHandle, x, DedRec
     vaSpread1.Col = 1
     vaSpread1.Row = x
     vaSpread1.Text = DedRec.DCDESC1
  Next x
  Close DedHandle
  
  Call MakeGLDescIdx
  
End Sub

Private Sub fpBegDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpBegDate_LostFocus()
  If CheckValDate(fpBegDate.Text) = False Then
    MsgBox "Please enter a valid date as a Beginning Date"
    fpBegDate.SetFocus
  End If
End Sub

Private Sub fpCheckDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpCheckDate_LostFocus()
  If CheckValDate(fpCheckDate.Text) = False Then
    MsgBox "Please enter a valid date as a Check Date"
    fpCheckDate.SetFocus
  End If

End Sub

Private Sub fpcurrTotDeds_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    vaSpread1.SetFocus
    vaSpread1.SetActiveCell 1, 1
    KeyCode = 0
  ElseIf KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 3
    fptxtEIC.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpGLNum_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyDelete Then
    fpGLNum(Index).Text = ""
    fpGLNum(Index).Action = ActionClearSearchBuffer
  End If
  If KeyCode = vbKeySpace Then
    fpGLNum(Index).ListDown = True
  End If
  If fpGLNum(Index).ListDown = False Then
    If Index = 1 Then
      If KeyCode = vbKeyLeft Then
        vaTabPro1.ActiveTab = 0
        fptxtHrsWkd.SetFocus
        KeyCode = 0
      ElseIf KeyCode = vbKeyUp Then
        fptxtGLAmt(4).SetFocus
        KeyCode = 0
      ElseIf KeyCode = vbKeyDown Then
        SendKeys "{Tab}"
        KeyCode = 0
      End If
    ElseIf Index <> 1 Then
      If KeyCode = vbKeyDown Then
        SendKeys "{Tab}"
        KeyCode = 0
      ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtCheckNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    vaTabPro1.ActiveTab = 0
    fptxtHrsWkd.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtCheckNum_LostFocus()
  If Len(QPTrim$(fptxtCheckNum.Text)) = 0 Then
    fptxtCheckNum.Text = "0"
  End If
End Sub

Private Sub fptxtCmpHrs_Change()
  If CheckFor2ManyDecimals(fptxtCmpHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtCmpHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtCmpHrs.Text) = "" Then fptxtCmpHrs.Text = "0.00"
'  fptxtCmpHrs.Text = QPTrim$(Using("#,##0.00", fptxtCmpHrs.Text))

End Sub

Private Sub fptxtCmpHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtEIC_Change()
  If QPTrim(fptxtEIC.Text) = "" Then fptxtEIC.Text = 0
  fptxtEIC.Text = Using("$###,##0.00", fptxtEIC.Text)

End Sub

Private Sub fptxtEIC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fptxtRG.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 2
    vaSpread1.SetFocus
    vaSpread1.SetActiveCell 1, 1
    KeyCode = 0
  End If

End Sub

Private Sub fptxtEIC_LostFocus()
  Call ReFigure
End Sub

Private Sub fptxtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtEnd_LostFocus()
  If CheckValDate(fptxtEnd.Text) = False Then
    MsgBox "Please enter a valid date as a End Date"
    fptxtEnd.SetFocus
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
   Dim Number As String
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
   For x = Nextx To 4
     For y = 1 To TotalAccts
       If DescBuff(y) <> 0 Then
         Get GLDHandle, DescBuff(y), GLDesc(1)
         Number = QPTrim$(GLDesc(1).Num)
         fpGLNum(Nextx).InsertRow = QPTrim$(GLDesc(1).Title) & " " & Chr$(9) & Number & Chr(9) & QPTrim(ReplaceString(GLDesc(1).Num, "-", ""))
       End If
     Next y
     Nextx = Nextx + 1
   Next x
   Close GLDHandle

End Sub

Private Sub ParseTrans2Manual(TransRec() As TransRecType, ManTrans() As ManualTransRecType)
  Dim cnt As Integer
  
  ManTrans(1).PDSTART = TransRec(1).PayPdStart
  ManTrans(1).PDEND = TransRec(1).PayPdEnd
  ManTrans(1).ChkDate = TransRec(1).CheckDate
  ManTrans(1).CheckNum = TransRec(1).CheckNum

  ManTrans(1).RegHrs = TransRec(1).RegHrsWork
  ManTrans(1).SICKHRS = TransRec(1).SickUsed
  ManTrans(1).VACHRS = TransRec(1).VacUsed
  ManTrans(1).COMPHRS = TransRec(1).CompUsed
  ManTrans(1).HOLHOURS = TransRec(1).HOLHOURS
  ManTrans(1).PERSHRS = TransRec(1).PerHours

  ManTrans(1).OTHRSPD = TransRec(1).OTHrsPaid

  ManTrans(1).RegWage = TransRec(1).TotRegWage
  ManTrans(1).OTWage = TransRec(1).TotOTWage

  ManTrans(1).DISTACT1 = TransRec(1).TDist(1).DAcct
  ManTrans(1).WAGEAMT1 = TransRec(1).TDist(1).DRWage

  ManTrans(1).DISTACT2 = TransRec(1).TDist(2).DAcct
  ManTrans(1).WAGEAMT2 = TransRec(1).TDist(2).DRWage

  ManTrans(1).DISTACT3 = TransRec(1).TDist(3).DAcct
  ManTrans(1).WAGEAMT3 = TransRec(1).TDist(3).DRWage
  ManTrans(1).DISTACT4 = TransRec(1).TDist(4).DAcct
  ManTrans(1).WAGEAMT4 = TransRec(1).TDist(4).DRWage

  ManTrans(1).GrossPay = TransRec(1).GrossPay

  ManTrans(1).FEDTAX = TransRec(1).FedTaxAmt
  ManTrans(1).STATAX = TransRec(1).StaTaxAmt
  ManTrans(1).SOCTAX = TransRec(1).SocTaxAmt
  ManTrans(1).MEDTAX = TransRec(1).MedTaxAmt
  ManTrans(1).RETAMT = TransRec(1).RetireAmt
  ManTrans(1).TOTTAX = TransRec(1).TotTaxAmt

'  For cnt = 1 To 50
  For cnt = 1 To DedRecCnt
    ManTrans(1).DAmt(cnt) = TransRec(1).DAmt(cnt)
  Next
  ManTrans(1).TOTDED = TransRec(1).TotDedAmt
  ManTrans(1).EIC = TransRec(1).EICAmt
  ManTrans(1).NetPay = TransRec(1).NetPay

  ManTrans(1).FedGross = TransRec(1).FedGrossPay
  ManTrans(1).STAGROSS = TransRec(1).StaGrossPay
  ManTrans(1).SocGross = TransRec(1).SocGrossPay
  ManTrans(1).MedGross = TransRec(1).MedGrossPay
  ManTrans(1).RETGROSS = TransRec(1).RetGrossPay

End Sub

Private Sub ParseManual2Scrn(TransRec() As TransRecType, ManTrans() As ManualTransRecType)
  Dim cnt As Integer
  Dim Image$, Image1$
  Dim Today As String * 10
  Dim DedsVal As Double
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$

  
  Image = "$###,##0.00"
  Image1 = "#,##0.00"
  fpBegDate.Text = MakeRegDate(ManTrans(1).PDSTART)
  If QPTrim$(fpBegDate.Text) = "12-31-1979" Then fpBegDate.Text = Today
  
  fptxtEnd.Text = MakeRegDate(ManTrans(1).PDEND)
  If QPTrim$(fptxtEnd.Text) = "12-31-1979" Then fptxtEnd.Text = Today
  
  fpCheckDate.Text = MakeRegDate(ManTrans(1).ChkDate)
  If QPTrim$(fpCheckDate.Text) = "12-31-1979" Then fpCheckDate.Text = Today

  fptxtCheckNum.Text = ManTrans(1).CheckNum

  fptxtHrsWkd.Text = QPTrim$(Using(Image1, TransRec(1).RegHrsWork))
  fptxtPerHrs.Text = QPTrim$(Using(Image1, ManTrans(1).PERSHRS))
  fptxtSckHrs.Text = QPTrim$(Using(Image1, ManTrans(1).SICKHRS))
  fptxtVacHrs.Text = QPTrim$(Using(Image1, ManTrans(1).VACHRS))
  fptxtCmpHrs.Text = QPTrim$(Using(Image1, ManTrans(1).COMPHRS))
  fptxtHolHrs.Text = QPTrim$(Using(Image1, ManTrans(1).HOLHOURS))
  fptxtOTHrs.Text = QPTrim$(Using(Image1, ManTrans(1).OTHRSPD))

  fptxtREGData.Text = ManTrans(1).RegWage
  fptxtOTData.Text = ManTrans(1).OTWage

  fpGLNum(1).Text = QPTrim$(ManTrans(1).DISTACT1)
  fptxtGLAmt(1).Text = ManTrans(1).WAGEAMT1

  fpGLNum(2).Text = QPTrim$(ManTrans(1).DISTACT2)
  fptxtGLAmt(2).Text = ManTrans(1).WAGEAMT2

  fpGLNum(3).Text = QPTrim$(ManTrans(1).DISTACT3)
  fptxtGLAmt(3).Text = ManTrans(1).WAGEAMT3

  fpGLNum(4).Text = QPTrim$(ManTrans(1).DISTACT4)
  fptxtGLAmt(4).Text = ManTrans(1).WAGEAMT4

  fptxtGPData.Text = ManTrans(1).GrossPay
  fptxtFTWHData.Text = ManTrans(1).FEDTAX
  fptxtSTWHData.Text = ManTrans(1).STATAX
  fptxtSSWHData.Text = ManTrans(1).SOCTAX
  fptxtMWHData.Text = ManTrans(1).MEDTAX
  fptxtRWHData.Text = ManTrans(1).RETAMT
  fptxtTotal.Text = ManTrans(1).FEDTAX + ManTrans(1).STATAX + ManTrans(1).SOCTAX + ManTrans(1).MEDTAX + ManTrans(1).RETAMT
  For cnt = 1 To DedRecCnt
    vaSpread1.Col = 2
    vaSpread1.Row = cnt
    If ManTrans(1).DAmt(cnt) > 0 Then
      vaSpread1.Text = QPTrim$(Using(Image, ManTrans(1).DAmt(cnt)))
      DedsVal = DedsVal + ManTrans(1).DAmt(cnt)
    Else
      vaSpread1.Text = QPTrim$(Using(Image, 0))
    End If
  Next
  fpcurrTotDeds.Text = Using(Image, DedsVal)
  fptxtEIC.Text = ManTrans(1).EIC
  fptxtNP.Text = ManTrans(1).NetPay
  fptxtFG.Text = ManTrans(1).FedGross
  fptxtStaG.Text = ManTrans(1).STAGROSS
  fptxtSocG.Text = ManTrans(1).SocGross
  fptxtMG.Text = ManTrans(1).MedGross
  fptxtRG.Text = ManTrans(1).RETGROSS
  
End Sub

Private Sub fptxtFG_Change()
  If QPTrim(fptxtFG.Text) = "" Then fptxtFG.Text = 0
  fptxtFG.Text = Using("$###,##0.00", fptxtFG.Text)

End Sub

Private Sub fptxtFG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtFTWHData_Change()
  If QPTrim(fptxtFTWHData.Text) = "" Then fptxtFTWHData.Text = 0
  fptxtFTWHData.Text = Using("$###,##0.00", fptxtFTWHData.Text)

End Sub

Private Sub fptxtFTWHData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtFTWHData_LostFocus()
  Call ReFigure

End Sub

Private Sub fptxtGLAmt_Change(Index As Integer)
  
  If QPTrim(fptxtGLAmt(Index).Text) = "" Then fptxtGLAmt(Index).Text = 0
End Sub

Private Sub fptxtGLAmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 4 Then
    If KeyCode = vbKeyRight Then
      vaTabPro1.ActiveTab = 2
      vaSpread1.SetFocus
      vaSpread1.SetActiveCell 1, 1
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      fpGLNum(1).SetFocus
      KeyCode = 0
    End If
  End If
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtGLAmt_LostFocus(Index As Integer)
  Call ReFigure
End Sub

Private Sub fptxtGPData_Change()
  If QPTrim(fptxtGPData.Text) = "" Then fptxtGPData.Text = 0
  fptxtGPData.Text = Using("$###,##0.00", fptxtGPData.Text)

End Sub

Private Sub fptxtGPData_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 3
    fptxtEIC.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fptxtTotal.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtHolHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtHolHrs_LostFocus()
  If CheckFor2ManyDecimals(fptxtHolHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtHolHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtHolHrs.Text) = "" Then fptxtHolHrs.Text = 0
  fptxtHolHrs.Text = Using("#,##0.00", fptxtHolHrs.Text)

End Sub

Private Sub fptxtHrsWkd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyLeft Then
    vaTabPro1.ActiveTab = 4
    fptxtGPData.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fptxtOTHrs.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub fptxtHrsWkd_LostFocus()
  If CheckFor2ManyDecimals(fptxtHrsWkd.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtHrsWkd.SetFocus
    Exit Sub
  End If
  
  If QPTrim(fptxtHrsWkd.Text) = "" Then fptxtHrsWkd.Text = 0
  fptxtHrsWkd.Text = Using("#,##0.00", fptxtHrsWkd.Text)

End Sub

Private Sub fptxtMG_Change()
  If QPTrim(fptxtMG.Text) = "" Then fptxtMG.Text = 0
  fptxtMG.Text = Using("$###,##0.00", fptxtMG.Text)

End Sub

Private Sub fptxtMG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtMWHData_Change()
  If QPTrim(fptxtMWHData.Text) = "" Then fptxtMWHData.Text = 0
  fptxtMWHData.Text = Using("$###,##0.00", fptxtMWHData.Text)

End Sub

Private Sub fptxtMWHData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtMWHData_LostFocus()
  Call ReFigure

End Sub

Private Sub fptxtNP_Change()
  If QPTrim(fptxtNP.Text) = "" Then fptxtNP.Text = 0
  fptxtNP.Text = Using("$###,##0.00", fptxtNP.Text)

End Sub

Private Sub fptxtNP_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtOTData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtOTData_LostFocus()
  Call ReFigure 'added 7/11/06

End Sub

Private Sub fptxtOTHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtHrsWkd.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 1
    fpGLNum(1).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtOTHrs_LostFocus()
  If CheckFor2ManyDecimals(fptxtOTHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtOTHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtOTHrs.Text) = "" Then fptxtOTHrs.Text = 0
  fptxtOTHrs.Text = Using("#,##0.00", fptxtOTHrs.Text)

End Sub

Private Sub fptxtPerHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtPerHrs_LostFocus()
  If CheckFor2ManyDecimals(fptxtPerHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtPerHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtPerHrs.Text) = "" Then fptxtPerHrs.Text = 0
  fptxtPerHrs.Text = Using("#,##0.00", fptxtPerHrs.Text)

End Sub

Private Sub fptxtREGData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fptxtRG_Change()
  If QPTrim(fptxtRG.Text) = "" Then fptxtRG.Text = 0
  fptxtRG.Text = Using("$###,##0.00", fptxtRG.Text)

End Sub

Private Sub fptxtRG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 4
    fptxtGPData.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fptxtEIC.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub fptxtRWHData_Change()
  If QPTrim(fptxtRWHData.Text) = "" Then fptxtRWHData.Text = 0
  fptxtRWHData.Text = Using("$###,##0.00", fptxtRWHData.Text)

End Sub

Private Sub fptxtRWHData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtRWHData_LostFocus()
  Call ReFigure

End Sub

Private Sub fptxtSckHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtSckHrs_LostFocus()
  If CheckFor2ManyDecimals(fptxtSckHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtSckHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtSckHrs.Text) = "" Then fptxtSckHrs.Text = 0
  fptxtSckHrs.Text = Using("#,##0.00", fptxtSckHrs.Text)

End Sub

Private Sub fptxtSocG_Change()
  If QPTrim(fptxtSocG.Text) = "" Then fptxtSocG.Text = 0
  fptxtSocG.Text = Using("$###,##0.00", fptxtSocG.Text)

End Sub

Private Sub fptxtSocG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtSSWHData_Change()
  If QPTrim(fptxtSSWHData.Text) = "" Then fptxtSSWHData.Text = 0
  fptxtSSWHData.Text = Using("$###,##0.00", fptxtSSWHData.Text)

End Sub

Private Sub fptxtSSWHData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtSSWHData_LostFocus()
  Call ReFigure

End Sub

Private Sub fptxtStaG_Change()
  If QPTrim(fptxtStaG.Text) = "" Then fptxtStaG.Text = 0
  fptxtStaG.Text = Using("$###,##0.00", fptxtStaG.Text)

End Sub

Private Sub fptxtStaG_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtSTWHData_Change()
  If QPTrim(fptxtSTWHData.Text) = "" Then fptxtSTWHData.Text = 0
  fptxtSTWHData.Text = Using("$###,##0.00", fptxtSTWHData.Text)

End Sub

Private Sub fptxtSTWHData_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtSTWHData_LostFocus()
  Call ReFigure
End Sub

Private Sub fptxtTotal_Change()
  If QPTrim(fptxtTotal.Text) = "" Then fptxtTotal.Text = 0
  fptxtTotal.Text = Using("$###,##0.00", fptxtTotal.Text)

End Sub

Private Sub fptxtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyRight Then
    vaTabPro1.ActiveTab = 0
    fptxtHrsWkd.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fptxtGPData.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub fptxtVacHrs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtVacHrs_LostFocus()
  If CheckFor2ManyDecimals(fptxtVacHrs.Text) = True Then
    MsgBox "Number entered is not valid"
    fptxtVacHrs.SetFocus
    Exit Sub
  End If
  If QPTrim(fptxtVacHrs.Text) = "" Then fptxtVacHrs.Text = 0
  fptxtVacHrs.Text = Using("#,##0.00", fptxtVacHrs.Text)

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
  If vaSpread1.Col = 2 Then
    vaSpread1.Row = Row
    If QPTrim$(vaSpread1.Text) = "" Then
      vaSpread1.Text = "$0.00"
      Exit Sub
    End If
    vaSpread1.Col = 2
    vaSpread1.Row = Row
    vaSpread1.Text = Format(vaSpread1.Text, "$###,##0.00")
  End If
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
  If (Col > 1) Then
    vaSpread1.EditMode = True
  Else
    vaSpread1.EditMode = False
  End If

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyPageUp Then
    vaTabPro1.ActiveTab = 3
    fptxtEIC.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyPageDown Then
    vaTabPro1.ActiveTab = 1
    fptxtGLAmt(4).SetFocus
    KeyCode = 0
  End If
  If Shift = 1 Then
    vaTabPro1.ActiveTab = 1
    fpGLNum(1).SetFocus
    KeyCode = 0
  End If
  
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Call ReFigure

End Sub

Private Sub ParseManual2Trans(TransRec() As TransRecType, ManTrans() As ManualTransRecType)
  Dim cnt As Integer
  TransRec(1).PayPdStart = ManTrans(1).PDSTART
  TransRec(1).PayPdEnd = ManTrans(1).PDEND

  TransRec(1).CheckDate = ManTrans(1).ChkDate

  TransRec(1).CheckNum = ManTrans(1).CheckNum
  'STOP

  TransRec(1).RegHrsWork = ManTrans(1).RegHrs
  
  TransRec(1).RegHrsPaid = ManTrans(1).RegHrs 'ManTrans(1).SICKHRS + ManTrans(1).VACHRS + ManTrans(1).COMPHRS + ManTrans(1).HOLHOURS
  
  TransRec(1).SickUsed = ManTrans(1).SICKHRS
  TransRec(1).VacUsed = ManTrans(1).VACHRS
  TransRec(1).CompUsed = ManTrans(1).COMPHRS
  TransRec(1).HOLHOURS = ManTrans(1).HOLHOURS
  TransRec(1).PerHours = ManTrans(1).PERSHRS
  TransRec(1).OTHrsPaid = ManTrans(1).OTHRSPD
  TransRec(1).OTHours = Val(fptxtOTHrs.Text) 'no ManTrans(1) field for this in type structure
  TransRec(1).TotRegWage = ManTrans(1).RegWage
  TransRec(1).TotOTWage = fptxtOTData.Text

  TransRec(1).TDist(1).DAcct = ManTrans(1).DISTACT1
  TransRec(1).TDist(1).DRWage = ManTrans(1).WAGEAMT1

  TransRec(1).TDist(2).DAcct = ManTrans(1).DISTACT2
  TransRec(1).TDist(2).DRWage = ManTrans(1).WAGEAMT2

  TransRec(1).TDist(3).DAcct = ManTrans(1).DISTACT3
  TransRec(1).TDist(3).DRWage = ManTrans(1).WAGEAMT3

  TransRec(1).TDist(4).DAcct = ManTrans(1).DISTACT4
  TransRec(1).TDist(4).DRWage = ManTrans(1).WAGEAMT4

  TransRec(1).GrossPay = ManTrans(1).GrossPay
  TransRec(1).FedTaxAmt = ManTrans(1).FEDTAX
  TransRec(1).StaTaxAmt = ManTrans(1).STATAX
  TransRec(1).SocTaxAmt = ManTrans(1).SOCTAX
  TransRec(1).MedTaxAmt = ManTrans(1).MEDTAX
  TransRec(1).RetireAmt = ManTrans(1).RETAMT
  TransRec(1).TotTaxAmt = ManTrans(1).TOTTAX

  For cnt = 1 To DedRecCnt
    TransRec(1).DAmt(cnt) = ManTrans(1).DAmt(cnt)
  Next

  TransRec(1).TotDedAmt = ManTrans(1).TOTDED
  TransRec(1).EICAmt = ManTrans(1).EIC

  TransRec(1).NetPay = ManTrans(1).NetPay

  TransRec(1).FedGrossPay = ManTrans(1).FedGross
  TransRec(1).StaGrossPay = ManTrans(1).STAGROSS
  TransRec(1).SocGrossPay = ManTrans(1).SocGross
  TransRec(1).MedGrossPay = ManTrans(1).MedGross
  TransRec(1).RetGrossPay = ManTrans(1).RETGROSS
  
  'these fields are not updated in manual processing
  'and therefore retain whatever value is carried over
  'it is best to zero these fields out so that (like
  'in the Retirement report using matching amounts) the
  'program is not using faulty data left over from the
  'last payroll....8/13/2003
  TransRec(1).MatchMedAmt = 0
  TransRec(1).MatchRetAmt = 0
  TransRec(1).MatchSocAmt = 0
    
End Sub

Private Sub ParseScrn2Manual(TransRec() As TransRecType, ManTrans() As ManualTransRecType)
  Dim cnt As Integer
  Dim Today As String * 10
  
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$

  ManTrans(1).PDSTART = Date2Num(fpBegDate.Text)
  
  ManTrans(1).PDEND = Date2Num(fptxtEnd.Text)
  
  ManTrans(1).ChkDate = Date2Num(fpCheckDate.Text)

  ManTrans(1).CheckNum = Val(fptxtCheckNum.Text)

  ManTrans(1).RegHrs = Val(fptxtHrsWkd.Text)
  ManTrans(1).PERSHRS = Val(fptxtPerHrs.Text)
  ManTrans(1).SICKHRS = Val(fptxtSckHrs.Text)
  ManTrans(1).VACHRS = Val(fptxtVacHrs.Text)
  ManTrans(1).COMPHRS = Val(fptxtCmpHrs.Text)
  ManTrans(1).HOLHOURS = Val(fptxtHolHrs.Text)
  ManTrans(1).OTHRSPD = Val(fptxtOTHrs.Text) 'number of hours

  ManTrans(1).RegWage = fptxtREGData.Text
  ManTrans(1).OTWage = fptxtOTData.Text 'dollars paid
  
  ManTrans(1).DISTACT1 = QPTrim$(fpGLNum(1).Text)
  ManTrans(1).WAGEAMT1 = fptxtGLAmt(1).Text

  ManTrans(1).DISTACT2 = QPTrim$(fpGLNum(2).Text)
  ManTrans(1).WAGEAMT2 = fptxtGLAmt(2).Text

  ManTrans(1).DISTACT3 = QPTrim$(fpGLNum(3).Text)
  ManTrans(1).WAGEAMT3 = fptxtGLAmt(3).Text

  ManTrans(1).DISTACT4 = QPTrim$(fpGLNum(4).Text)
  ManTrans(1).WAGEAMT4 = fptxtGLAmt(4).Text

  ManTrans(1).GrossPay = fptxtGPData.Text
  ManTrans(1).FEDTAX = fptxtFTWHData.Text
  ManTrans(1).STATAX = fptxtSTWHData.Text
  ManTrans(1).SOCTAX = fptxtSSWHData.Text
  ManTrans(1).MEDTAX = fptxtMWHData.Text
  ManTrans(1).RETAMT = fptxtRWHData.Text
  'changed next line on 10/3/06
  ManTrans(1).TOTTAX = ManTrans(1).FEDTAX + ManTrans(1).STATAX + ManTrans(1).SOCTAX + ManTrans(1).MEDTAX + ManTrans(1).RETAMT
  
  For cnt = 1 To DedRecCnt
    vaSpread1.Col = 2
    vaSpread1.Row = cnt
    ManTrans(1).DAmt(cnt) = Val(ReplaceString$(vaSpread1.Text, "$", "")) 'RemoveDollarMark
  Next

  ManTrans(1).TOTDED = fpcurrTotDeds.Text
  ManTrans(1).EIC = fptxtEIC.Text

  ManTrans(1).NetPay = fptxtNP.Text

  ManTrans(1).FedGross = fptxtFG.Text
  ManTrans(1).STAGROSS = fptxtStaG.Text
  ManTrans(1).SocGross = fptxtSocG.Text
  ManTrans(1).MedGross = fptxtMG.Text
  ManTrans(1).RETGROSS = fptxtRG.Text
  
End Sub

Private Sub ReFigure()
  Dim Image$
  Dim AllDeds As Double
  Dim NetPay As Double
  Dim AllWH As Double
  Dim x As Integer
  Dim AllGLDeds As Double
  
  AllGLDeds = 0
  AllWH = 0
  AllDeds = 0
  Image = "$###,##0.00"
  
  For x = 1 To 4
    AllGLDeds = AllGLDeds + CDbl(fptxtGLAmt(x).Text)
  Next x
  
  fptxtGPData.Text = Format(AllGLDeds, Image)
  fptxtREGData.Text = Format(AllGLDeds, Image) - CDbl(fptxtOTData.Text) '7/11/06 added OT
'  fptxtFG.Text = Format(AllGLDeds, Image)
'  fptxtStaG.Text = Format(AllGLDeds, Image)
'  fptxtSocG.Text = Format(AllGLDeds, Image)
'  fptxtMG.Text = Format(AllGLDeds, Image)
'  fptxtRG.Text = Format(AllGLDeds, Image)
  
  AllWH = CDbl(fptxtFTWHData) + CDbl(fptxtSTWHData)
  AllWH = AllWH + CDbl(fptxtMWHData.Text) + CDbl(fptxtSSWHData.Text)
  AllWH = AllWH + CDbl(fptxtRWHData.Text)
  
  For x = 1 To DedRecCnt
    vaSpread1.Col = 2
    vaSpread1.Row = x
    AllDeds = AllDeds + Val(ReplaceString$(vaSpread1.Text, "$", "")) 'RemoveDollarMark
  Next x
  
  fpcurrTotDeds.Text = AllDeds
  NetPay = AllGLDeds - (AllWH + AllDeds) + CDbl(fptxtEIC.Text)
  fptxtNP.Text = Format(NetPay, Image)
  fptxtTotal.Text = Format(AllWH, Image) ' - CDbl(fptxtRWHData.Text) 'took out RWH on 7/11/06

End Sub

Private Sub LoadScrnTActiveIsFalse()
  Dim cnt As Integer
  Dim Emp2Rec As EmpData2Type
  Dim E2Handle As Integer
  
  Dim Image$
  Dim Today As String * 10
  Dim DedsVal As Double
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$

  Image = "$###,##0.00"
  
  OpenEmpData2File E2Handle
  Get E2Handle, RecNum, Emp2Rec
  Close E2Handle
  fpBegDate.Text = Today
  
  fptxtEnd.Text = Today
  
  fpCheckDate.Text = Today

  fptxtCheckNum.Text = "0"

  fptxtHrsWkd.Text = "0.00"
  fptxtPerHrs.Text = "0.00"
  fptxtSckHrs.Text = "0.00"
  fptxtVacHrs.Text = "0.00"
  fptxtCmpHrs.Text = "0.00"
  fptxtHolHrs.Text = "0.00"
  fptxtOTHrs.Text = "0.00"

  fptxtREGData.Text = QPTrim$(Using(Image, 0))
  fptxtOTData.Text = QPTrim$(Using(Image, 0))

  fpGLNum(1).Text = QPTrim$(Emp2Rec.EDist(1).DAcct)
  fptxtGLAmt(1).Text = QPTrim$(Using(Image, 0))

  fpGLNum(2).Text = QPTrim$(Emp2Rec.EDist(2).DAcct)
  fptxtGLAmt(2).Text = QPTrim$(Using(Image, 0))

  fpGLNum(3).Text = QPTrim$(Emp2Rec.EDist(3).DAcct)
  fptxtGLAmt(3).Text = QPTrim$(Using(Image, 0))

  fpGLNum(4).Text = QPTrim$(Emp2Rec.EDist(4).DAcct)
  fptxtGLAmt(4).Text = QPTrim$(Using(Image, 0))

  fptxtGPData.Text = QPTrim$(Using(Image, 0))
  fptxtFTWHData.Text = QPTrim$(Using(Image, 0))
  fptxtSTWHData.Text = QPTrim$(Using(Image, 0))
  fptxtSSWHData.Text = QPTrim$(Using(Image, 0))
  fptxtMWHData.Text = QPTrim$(Using(Image, 0))
  fptxtRWHData.Text = QPTrim$(Using(Image, 0))
  fptxtTotal.Text = QPTrim$(Using(Image, 0))
  For cnt = 1 To DedRecCnt
    vaSpread1.Col = 2
    vaSpread1.Row = cnt
    vaSpread1.Text = QPTrim$(Using(Image, 0))
  Next

  fpcurrTotDeds.Text = Using(Image, 0)
  fptxtEIC.Text = QPTrim$(Using(Image, 0))

  fptxtNP.Text = QPTrim$(Using(Image, 0))

  fptxtFG.Text = QPTrim$(Using(Image, 0))
  fptxtStaG.Text = QPTrim$(Using(Image, 0))
  fptxtSocG.Text = QPTrim$(Using(Image, 0))
  fptxtMG.Text = QPTrim$(Using(Image, 0))
  fptxtRG.Text = QPTrim$(Using(Image, 0))
  
End Sub

Private Sub DeleteThisEmp()
  Dim THandle As Integer
  Dim TransRec As TransRecType
  
  OpenTransWorkFile THandle
  Get THandle, RecNum, TransRec
  TransRec.TActive = 0
  Put THandle, RecNum, TransRec
  
  Close THandle

End Sub

Private Sub vaSpread1_LostFocus()
  Dim x As Integer
  
  For x = 1 To DedRecCnt
    vaSpread1.Col = 2
    vaSpread1.Row = x
    If CheckFor2ManyDecimals(vaSpread1.Text) = True Then
      MsgBox "Please enter a valid number"
      vaSpread1.SetFocus
      vaSpread1.SetActiveCell 2, x
      Exit Sub
    End If
  Next x
End Sub
Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
  Select Case ScreenW
    Case 1280
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 20
      coladj = 10
      vaTabPro1.TabHeight = 500
      For cnt = 1 To 5
        vaTabPro1.Tab = cnt - 1
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 16
        vaTabPro1.FontBold = False
      Next cnt
      vaSpread1.FontSize = 16
      vaSpread1.RowHeight(-1) = 22
      vaSpread1.RowHeight(0) = 22
    Else
      COne = 10.5
      coladj = 3.5
      vaSpread1.RowHeight(-1) = 18
      vaSpread1.RowHeight(0) = 18
    End If
    Case 1152
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 15
      coladj = 6.75
      vaTabPro1.TabHeight = 450
      vaSpread1.FontSize = 14
      vaSpread1.RowHeight(0) = 18.5
      vaSpread1.RowHeight(-1) = 18.5
    Else
      COne = 3.75
      coladj = 2.75
      For cnt = 1 To 5
        vaTabPro1.Tab = cnt - 1
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 12
        vaTabPro1.FontBold = False
      Next cnt
      vaSpread1.RowHeight(0) = 15
      vaSpread1.RowHeight(-1) = 15
    End If
    Case 1024
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 9.25
      coladj = 4.2
      vaTabPro1.TabHeight = 450
      vaTabPro1.FontBold = False '7/22
      vaSpread1.FontBold = True
      vaSpread1.RowHeight(0) = 17.5
      vaSpread1.RowHeight(-1) = 17.5
    Else
      COne = 0
      coladj = 0
    End If
    Case 800
      COne = 0
      coladj = -1.5
      vaSpread1.Font.Size = 10
      vaSpread1.RowHeight(-1) = 12.2
      For cnt = 1 To 5
        vaTabPro1.Tab = cnt - 1
        vaTabPro1.FontName = "Tahoma"
        vaTabPro1.FontSize = 10
        vaTabPro1.FontBold = True
      Next cnt
      vaTabPro1.ActiveTabBold = False
    Case Else
     
  End Select
  vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1) + COne
  vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + coladj

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmManualTransEntry.")
      Call Terminate
      End
    End If
  End If
End Sub

