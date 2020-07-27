VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxSystemSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax System Setup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "frmTaxSystemSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   990
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7920
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
      ButtonDesigner  =   "frmTaxSystemSetup.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   630
      Left            =   8535
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7920
      Width           =   2385
      _Version        =   131072
      _ExtentX        =   4207
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmTaxSystemSetup.frx":0AA9
   End
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6375
      Left            =   240
      TabIndex        =   32
      Top             =   1440
      Width           =   11535
      _Version        =   196609
      _ExtentX        =   20346
      _ExtentY        =   11245
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
      TabCaption      =   "frmTaxSystemSetup.frx":0C86
      PageEarMarkPictureNext=   "frmTaxSystemSetup.frx":0EFA
      PageEarMarkPicturePrev=   "frmTaxSystemSetup.frx":0F16
      EarMarkPictureNext=   "frmTaxSystemSetup.frx":0F32
      EarMarkPicturePrev=   "frmTaxSystemSetup.frx":0F4E
      Begin ImpproLib.vaImprint vaImprint6 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   33
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":0F6A
         Begin EditLib.fpText fptxtAN 
            Height          =   384
            Index           =   1
            Left            =   2880
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   36
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   46
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
            TabIndex        =   47
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
            TabIndex        =   48
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   1104
            Width           =   1308
         End
      End
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   52
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":0F86
         Begin LpLib.fpCombo fpcomboFedX 
            Height          =   405
            Left            =   1935
            TabIndex        =   61
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
            ColDesigner     =   "frmTaxSystemSetup.frx":0FA2
         End
         Begin LpLib.fpCombo fpcomboStateX 
            Height          =   405
            Left            =   1935
            TabIndex        =   60
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
            ColDesigner     =   "frmTaxSystemSetup.frx":13ED
         End
         Begin LpLib.fpCombo fpcomboFedAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   59
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
            ColDesigner     =   "frmTaxSystemSetup.frx":1838
         End
         Begin LpLib.fpCombo fpcomboStateAmtPct 
            Height          =   405
            Left            =   3075
            TabIndex        =   58
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
            ColDesigner     =   "frmTaxSystemSetup.frx":1C83
         End
         Begin LpLib.fpCombo fpcomboStateStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   57
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
            ColDesigner     =   "frmTaxSystemSetup.frx":20CE
         End
         Begin LpLib.fpCombo fpcomboFedStatus 
            Height          =   405
            Left            =   5670
            TabIndex        =   56
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
            ColDesigner     =   "frmTaxSystemSetup.frx":2571
         End
         Begin LpLib.fpCombo fpcomboSocX 
            Height          =   405
            Left            =   6330
            TabIndex        =   55
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
            ColDesigner     =   "frmTaxSystemSetup.frx":2A14
         End
         Begin LpLib.fpCombo fpcomboMedX 
            Height          =   405
            Left            =   6000
            TabIndex        =   54
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
            ColDesigner     =   "frmTaxSystemSetup.frx":2E5F
         End
         Begin LpLib.fpCombo fpcomboEIC 
            Height          =   405
            Left            =   4710
            TabIndex        =   53
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
            ColDesigner     =   "frmTaxSystemSetup.frx":32AA
         End
         Begin EditLib.fpCurrency fptxtAddWHFed 
            Height          =   396
            Left            =   8352
            TabIndex        =   62
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
            TabIndex        =   63
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
            TabIndex        =   64
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
            TabIndex        =   65
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
            TabIndex        =   66
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
            TabIndex        =   67
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
         TabIndex        =   81
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":374D
         Begin LpLib.fpCombo fpcomboFreq 
            Height          =   405
            Left            =   2685
            TabIndex        =   84
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
            ColDesigner     =   "frmTaxSystemSetup.frx":3769
         End
         Begin LpLib.fpCombo fpcomboStatus 
            Height          =   405
            Left            =   2685
            TabIndex        =   83
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
            ColDesigner     =   "frmTaxSystemSetup.frx":3BB4
         End
         Begin LpLib.fpCombo fpcomboPayType 
            Height          =   405
            Left            =   2685
            TabIndex        =   82
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
            ColDesigner     =   "frmTaxSystemSetup.frx":3FFF
         End
         Begin EditLib.fpCurrency fptxtRate 
            Height          =   450
            Left            =   7830
            TabIndex        =   85
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
            TabIndex        =   86
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
            TabIndex        =   87
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
            TabIndex        =   88
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
            TabIndex        =   89
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
            TabIndex        =   90
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
            TabIndex        =   91
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
            TabIndex        =   92
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
            TabIndex        =   93
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
            Top             =   645
            Width           =   690
         End
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   5865
         Left            =   150
         TabIndex        =   106
         Top             =   390
         Width           =   11265
         _Version        =   196609
         _ExtentX        =   19870
         _ExtentY        =   10345
         _StockProps     =   70
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":444A
         Begin LpLib.fpCombo fpcmbStateOfTax 
            Height          =   390
            Left            =   7635
            TabIndex        =   6
            Tag             =   $"frmTaxSystemSetup.frx":4466
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
            ColDesigner     =   "frmTaxSystemSetup.frx":45CE
         End
         Begin LpLib.fpCombo fpcmbMinOptions 
            Height          =   330
            Left            =   6360
            TabIndex        =   14
            Tag             =   $"frmTaxSystemSetup.frx":4A19
            Top             =   5280
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
            ColDesigner     =   "frmTaxSystemSetup.frx":4B81
         End
         Begin LpLib.fpList fpListTownships 
            Height          =   300
            Left            =   8760
            TabIndex        =   198
            Top             =   1560
            Width           =   2295
            _Version        =   196608
            _ExtentX        =   4048
            _ExtentY        =   529
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
            ColDesigner     =   "frmTaxSystemSetup.frx":4FCC
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdAddTownship 
            Height          =   375
            Left            =   8760
            TabIndex        =   201
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":53AC
         End
         Begin EditLib.fpCurrency fptxtMinTaxAmt 
            Height          =   372
            Left            =   2640
            TabIndex        =   13
            Top             =   5280
            Width           =   1092
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
         Begin EditLib.fpDoubleSingle fptxtCurrYrIntRate 
            Height          =   372
            Left            =   3240
            TabIndex        =   7
            ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
            Top             =   2772
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
            Left            =   9840
            TabIndex        =   9
            ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
            Top             =   2772
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
            Left            =   1680
            TabIndex        =   10
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   3570
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
            TabIndex        =   11
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   3570
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
         Begin EditLib.fpDateTime fptxtCurrYear 
            Height          =   345
            Left            =   6240
            TabIndex        =   8
            Top             =   2775
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
            TabIndex        =   199
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
         Begin EditLib.fpText fptxtPersOptSrch 
            Height          =   390
            Left            =   8880
            TabIndex        =   12
            Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
            Top             =   3570
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
         Begin VB.Label Label71 
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
            TabIndex        =   215
            Top             =   3690
            Width           =   1380
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   11160
            Y1              =   3240
            Y2              =   3240
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
            TabIndex        =   212
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Line Line7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   8640
            X2              =   8640
            Y1              =   120
            Y2              =   2520
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
            TabIndex        =   200
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   197
            Top             =   2880
            Width           =   1740
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
            TabIndex        =   195
            Top             =   3690
            Width           =   1260
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
            Left            =   240
            TabIndex        =   194
            Top             =   3690
            Width           =   1380
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
            Height          =   270
            Left            =   120
            TabIndex        =   193
            Top             =   3240
            Width           =   2700
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
            TabIndex        =   191
            Top             =   120
            Width           =   2340
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   11160
            Y1              =   2520
            Y2              =   2520
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
            Height          =   270
            Left            =   120
            TabIndex        =   190
            Top             =   2520
            Width           =   3060
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   2415
            Left            =   120
            Top             =   120
            Width           =   11055
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   1545
            Left            =   120
            Top             =   2520
            Width           =   11055
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
            TabIndex        =   119
            Top             =   435
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
            TabIndex        =   118
            Top             =   825
            Width           =   2460
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
            TabIndex        =   117
            Top             =   1230
            Width           =   2460
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
            TabIndex        =   116
            Top             =   1605
            Width           =   900
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
            TabIndex        =   115
            Top             =   1995
            Width           =   1020
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
            TabIndex        =   114
            Top             =   1995
            Width           =   1020
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
            TabIndex        =   113
            Top             =   1995
            Width           =   1500
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Current Year Interest Rate:"
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
            Left            =   600
            TabIndex        =   112
            Top             =   2880
            Width           =   2580
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
            Left            =   7440
            TabIndex        =   111
            Top             =   2880
            Width           =   2340
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmTaxSystemSetup.frx":5590
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
            Height          =   852
            Left            =   240
            TabIndex        =   110
            Top             =   4392
            Width           =   10620
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
            Height          =   276
            Left            =   360
            TabIndex        =   109
            Top             =   5364
            Width           =   2220
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
            Height          =   276
            Left            =   4080
            TabIndex        =   108
            Top             =   5364
            Width           =   2220
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
            TabIndex        =   107
            Top             =   4065
            Width           =   2340
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   1695
            Index           =   0
            Left            =   120
            Top             =   4050
            Width           =   11055
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   5865
         Left            =   -26310
         TabIndex        =   120
         Top             =   -21225
         Width           =   11265
         _Version        =   196609
         _ExtentX        =   19870
         _ExtentY        =   10345
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":56E6
         Begin LpLib.fpCombo cmbAutoFill 
            Height          =   390
            Left            =   10275
            TabIndex        =   216
            Tag             =   $"frmTaxSystemSetup.frx":5702
            ToolTipText     =   $"frmTaxSystemSetup.frx":586A
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
            ColDesigner     =   "frmTaxSystemSetup.frx":5913
         End
         Begin LpLib.fpCombo fpcmbRealPersSplitYN 
            Height          =   390
            Left            =   10320
            TabIndex        =   25
            Tag             =   $"frmTaxSystemSetup.frx":5D5E
            Top             =   3795
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
            ColDesigner     =   "frmTaxSystemSetup.frx":5EC6
         End
         Begin LpLib.fpCombo fpcmbCountyYN 
            Height          =   390
            Left            =   10320
            TabIndex        =   24
            Tag             =   $"frmTaxSystemSetup.frx":6311
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
            ColDesigner     =   "frmTaxSystemSetup.frx":6479
         End
         Begin LpLib.fpCombo fpcmbCyclesYN 
            Height          =   390
            Left            =   10320
            TabIndex        =   23
            Tag             =   $"frmTaxSystemSetup.frx":68C4
            Top             =   2760
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
            ColDesigner     =   "frmTaxSystemSetup.frx":6A2C
         End
         Begin LpLib.fpCombo fpcmbCentDepYN 
            Height          =   390
            Left            =   2640
            TabIndex        =   15
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
            ColDesigner     =   "frmTaxSystemSetup.frx":6E77
         End
         Begin LpLib.fpCombo fpcmbTaxBillFormat 
            Height          =   375
            Left            =   2115
            TabIndex        =   18
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
            ColDesigner     =   "frmTaxSystemSetup.frx":72C2
         End
         Begin LpLib.fpCombo fpcmbAcctMeth 
            Height          =   390
            Left            =   8715
            TabIndex        =   19
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
            ColDesigner     =   "frmTaxSystemSetup.frx":770D
         End
         Begin LpLib.fpCombo fpcmbLateFormat 
            Height          =   375
            Left            =   2115
            TabIndex        =   20
            Tag             =   $"frmTaxSystemSetup.frx":7B58
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
            ColDesigner     =   "frmTaxSystemSetup.frx":7CC0
         End
         Begin LpLib.fpCombo fpcmbNoInterYN 
            Height          =   390
            Left            =   10275
            TabIndex        =   21
            Tag             =   $"frmTaxSystemSetup.frx":810B
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
            ColDesigner     =   "frmTaxSystemSetup.frx":8273
         End
         Begin EditLib.fpDoubleSingle fptxtDiscPct 
            Height          =   372
            Left            =   9120
            TabIndex        =   26
            ToolTipText     =   $"frmTaxSystemSetup.frx":86BE
            Top             =   4368
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
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2895
            Left            =   240
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2640
            Width           =   5895
            _Version        =   196613
            _ExtentX        =   10398
            _ExtentY        =   5106
            _StockProps     =   64
            ArrowsExitEditMode=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            MaxCols         =   4
            MaxRows         =   7
            SpreadDesigner  =   "frmTaxSystemSetup.frx":8749
         End
         Begin EditLib.fpText fptxtOverPayGL 
            Height          =   390
            Left            =   3000
            TabIndex        =   22
            Top             =   1800
            Width           =   2175
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
            TabIndex        =   187
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":8C17
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdLateBill 
            Height          =   330
            Left            =   4800
            TabIndex        =   188
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":8DFB
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdCycle 
            Height          =   372
            Left            =   6360
            TabIndex        =   203
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1548
            _Version        =   131072
            _ExtentX        =   2730
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":8FE0
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdCounty 
            Height          =   372
            Left            =   6360
            TabIndex        =   205
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1548
            _Version        =   131072
            _ExtentX        =   2730
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":91C3
         End
         Begin EditLib.fpText fptxtCentCash 
            Height          =   390
            Left            =   5160
            TabIndex        =   16
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
            TabIndex        =   17
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
            Left            =   7080
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   5280
            Width           =   3108
            _Version        =   131072
            _ExtentX        =   5482
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
            ButtonDesigner  =   "frmTaxSystemSetup.frx":93A2
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Auto Fill Service Address:"
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
            Height          =   270
            Left            =   6840
            TabIndex        =   217
            Top             =   1890
            Width           =   3300
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Real Classification"
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
            Left            =   6240
            TabIndex        =   213
            Top             =   4920
            Width           =   2364
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   11160
            X2              =   6240
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            X1              =   6240
            X2              =   6240
            Y1              =   2280
            Y2              =   5730
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
            TabIndex        =   210
            Top             =   480
            Width           =   1380
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
            TabIndex        =   209
            Top             =   480
            Width           =   1500
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
            Left            =   10332
            TabIndex        =   208
            Top             =   4440
            Width           =   300
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Use Real/Personal Split Billing Y/N?:"
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
            Left            =   6600
            TabIndex        =   207
            Top             =   3888
            Width           =   3660
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
            Left            =   8040
            TabIndex        =   206
            Top             =   3360
            Width           =   2220
         End
         Begin VB.Shape Shape12 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   3465
            Left            =   120
            Top             =   2280
            Width           =   11055
         End
         Begin VB.Label Label82 
            Alignment       =   2  'Center
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
            Left            =   8040
            TabIndex        =   204
            Top             =   2880
            Width           =   2220
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
            Height          =   270
            Left            =   6240
            TabIndex        =   202
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   2
            Height          =   2175
            Left            =   120
            Top             =   120
            Width           =   11055
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
            TabIndex        =   196
            Top             =   1428
            Width           =   3300
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Principal Discount Pct:"
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
            TabIndex        =   192
            Top             =   4476
            Width           =   2148
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "Pay Sequence and Opt Revenue Setup"
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
            TabIndex        =   189
            Top             =   2280
            Width           =   4545
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Overpayment G/L Number:"
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
            TabIndex        =   186
            Top             =   1920
            Width           =   2580
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
            TabIndex        =   185
            Top             =   120
            Width           =   1545
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
            TabIndex        =   184
            Top             =   1425
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
            TabIndex        =   183
            Top             =   945
            Width           =   1740
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
            TabIndex        =   182
            Top             =   945
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
            TabIndex        =   181
            Top             =   468
            Width           =   2340
         End
      End
      Begin ImpproLib.vaImprint vaImprint5 
         Height          =   5025
         Left            =   -25350
         TabIndex        =   121
         Top             =   -20640
         Width           =   10305
         _Version        =   196609
         _ExtentX        =   18177
         _ExtentY        =   8864
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":9597
      End
      Begin ImpproLib.vaImprint vaImprint7 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   122
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":95B3
         Begin EditLib.fpText fptxtWDDD 
            Height          =   396
            Index           =   8
            Left            =   7776
            TabIndex        =   123
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
            TabIndex        =   124
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
            TabIndex        =   125
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
            TabIndex        =   126
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
            TabIndex        =   127
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
            TabIndex        =   128
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
            TabIndex        =   129
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
            TabIndex        =   130
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
            TabIndex        =   131
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
            TabIndex        =   132
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
            TabIndex        =   133
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
            TabIndex        =   134
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
            TabIndex        =   135
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
            TabIndex        =   136
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
            TabIndex        =   137
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
            TabIndex        =   138
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
            TabIndex        =   149
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
            TabIndex        =   148
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
            TabIndex        =   147
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
            TabIndex        =   146
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   139
            Top             =   384
            Width           =   1980
         End
      End
      Begin ImpproLib.vaImprint vaImprint8 
         Height          =   5100
         Left            =   -25290
         TabIndex        =   150
         Top             =   -20715
         Width           =   10245
         _Version        =   196609
         _ExtentX        =   18071
         _ExtentY        =   8996
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   9405029
         Caption         =   ""
         Picture         =   "frmTaxSystemSetup.frx":95CF
         Begin LpLib.fpCombo fpcomboLT 
            Height          =   405
            Left            =   2070
            TabIndex        =   153
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
            ColDesigner     =   "frmTaxSystemSetup.frx":95EB
         End
         Begin LpLib.fpCombo fpcomboESC 
            Height          =   405
            Left            =   9030
            TabIndex        =   152
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
            ColDesigner     =   "frmTaxSystemSetup.frx":9A36
         End
         Begin LpLib.fpCombo fpcombo401K 
            Height          =   405
            Left            =   5370
            TabIndex        =   151
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
            ColDesigner     =   "frmTaxSystemSetup.frx":9E81
         End
         Begin EditLib.fpText fptxtEarned 
            Height          =   348
            Index           =   2
            Left            =   3264
            TabIndex        =   154
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
            TabIndex        =   155
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
            TabIndex        =   156
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
            TabIndex        =   157
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
            TabIndex        =   158
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
            TabIndex        =   159
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
            TabIndex        =   160
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
            TabIndex        =   161
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
            TabIndex        =   162
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
            TabIndex        =   163
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
            TabIndex        =   164
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
            TabIndex        =   165
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
            TabIndex        =   166
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
            TabIndex        =   167
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
            TabIndex        =   168
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
            TabIndex        =   179
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
   Begin fpBtnAtlLibCtl.fpBtn cmdNextTab 
      Height          =   630
      Left            =   6015
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   7920
      Width           =   2400
      _Version        =   131072
      _ExtentX        =   4233
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmTaxSystemSetup.frx":A2CC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   636
      Left            =   3506
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   7920
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
      ButtonDesigner  =   "frmTaxSystemSetup.frx":A4AC
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
      Height          =   300
      Left            =   240
      TabIndex        =   31
      Top             =   1125
      Width           =   2460
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1493
      Top             =   350
      Width           =   8655
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
      Height          =   390
      Left            =   3143
      TabIndex        =   28
      Top             =   510
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1493
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxSystemSetup"
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
  Dim TempCurrYrInt As Double
  Dim TempPastYrInt As Double
  Dim TempPenPct As Double
  Dim StrEmpty As Boolean
  Dim TempTaxForm As Integer
  Dim TempMinTxOpt As String
  Dim TempMinTxPct As Double
  Dim TempAcctgMethod As String
  Dim TempDisPct As String
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
  Dim TempPenOpt1YN As String * 1
  Dim TempPenOpt2YN As String * 1
  Dim TempPenOpt3YN As String * 1
  Dim TempIntPrncTaxYN As String * 1
  Dim TempIntIntYN As String * 1
  Dim TempIntAdvYN As String * 1
  Dim TempIntLateLstYN As String * 1
  Dim TempIntOpt1YN As String * 1
  Dim TempIntOpt2YN As String * 1
  Dim TempIntOpt3YN As String * 1
  Dim TempOptRev1 As String * 35
  Dim TempOptRev2 As String * 35
  Dim TempOptRev3 As String * 35
  Dim TempDisStopDate As Integer
  Dim TempOptSrchCust As String
  Dim TempOptSrchProp As String
  Dim TempOptSrchPers As String
  Dim TempWarnInt As String * 1
  Dim TempTaxYear As Integer
  Dim TempUseCyclesYN As String
  Dim TempUseCountyYN As String
  Dim TempRealPersSplit As String
  Dim TempSnrCtzAmt As Double
  Dim SaveFlag As Boolean
  Dim TSListIdx As Integer
  Dim Fund As Integer, Dept As Integer, Detail As Integer
  
Private Sub cmdAddTownship_Click()
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Integer
  Dim ThisName$
  
  'on error goto ERRORSTUFF
  
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "cmdAddTownship_Click", Erl)
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
  frmTaxCountySetup.Show vbModal
End Sub

Private Sub cmdCycle_Click()
  frmTaxCycleSetup.Show vbModal
End Sub

Private Sub cmdExit_Click()

  'on error goto ERRORSTUFF
  Unload frmTaxGLList
  If Exist("TaxSetup.Dat") Then
    If Check4Changes = True Then
      Exit Sub
    End If
  Else
    frmTaxMsgWOpts.Label1.Caption = "You are exiting without saving any data. If you wish to continue exiting without saving then press F10. If you wish to return to the screen to save this data then press ESC and you will be returned to the screen where you can press the save button to save your data."
    frmTaxMsgWOpts.Label1.Top = 600
    frmTaxMsgWOpts.cmdCont.Text = "F10 OK To Exit"
    frmTaxMsgWOpts.cmdExit.Text = "ESC Abort Exit"
    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      GoTo ExitWOSaving
    Else
      Unload frmTaxMsgWOpts
      vaTabPro1.ActiveTab = 0
      fptxtNameOfTaxAuth.SetFocus
      Exit Sub
    End If
  End If
  
  Call LogSaves
  
ExitWOSaving:
  If Exist("C:\CPWork\lateltr.dat") Then
    KillFile "C:\CPWork\lateltr.dat"
    frmTaxBillingMenu.Show
    DoEvents
    Unload Me
  Else
    frmTaxBillSetUpMenu.Show
    DoEvents
    Unload Me
  End If
  MainLog ("User closed frmTaxSystemSetup.")
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "cmdExit_Click", Erl)
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
  frmTaxGLList.Show ' vbModal
End Sub

Private Sub cmdLateBill_Click()
  Dim x As Integer
  If fpcmbLateFormat.Text = "1) SELF EDIT #1" Then
    frmTaxLateNoticeLtr.Show
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

Private Sub cmdRealClass_Click()
  frmTaxRealClassSetup.Show vbModal
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
  
'  'on error goto ERRORSTUFF
  EditFlag = False
  SaveFlag = True
  If fpcmbAcctMeth.Text <> "NONE" Then
    If QPTrim$(fptxtOverPayGL.Text) = "" Then
      Call TaxMsg(800, "Please enter a valid General Ledger number in the 'Overpayment GL Number' field.")
      vaTabPro1.ActiveTab = 1
      fptxtOverPayGL.SetFocus
      Exit Sub
    End If
  End If
  
  If VerifyGLNum(QPTrim$(fptxtOverPayGL.Text)) = False Then
    frmTaxMsgWOpts.Label1.Caption = "The Overpayment GL number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
    frmTaxMsgWOpts.Label1.Top = 800
    frmTaxMsgWOpts.cmdCont.Text = "F10 Continue"

    frmTaxMsgWOpts.Show vbModal
    If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmTaxMsgWOpts
      MainLog ("Warning: User issued warning that the overpayment GL number " + QPTrim$(fptxtOverPayGL.Text) + " could not be verified and they elected to continue to save it anyway.")
    Else
      Unload frmTaxMsgWOpts
      Close
      vaTabPro1.ActiveTab = 1
      fptxtOverPayGL.SetFocus
      Exit Sub
    End If
  End If
    
  If QPTrim$(fptxtNameOfTaxAuth.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'Name of Taxing Authority'."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtNameOfTaxAuth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdd1.Text) = "" And QPTrim$(fptxtAdd2.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter an 'Address'."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'City' name."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtState.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'State' abbreviation for the state in which the town is located."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(ReplaceString(fptxtZip.Text, "-", "")) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'Zip Code' for the town."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbStateOfTax.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'State' abbreviation for the state in which the taxes will be paid."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbStateOfTax.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbAcctMeth.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter an 'Accounting Method'."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbAcctMeth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbTaxBillFormat.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'Tax Bill Format'."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbTaxBillFormat.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbLateFormat.Text) = "" Then
    frmTaxMsg.Label1.Caption = "Please enter a 'Late Bill Format'."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbLateFormat.SetFocus
    Exit Sub
  End If
  
  PenCnt = 0
  For x = 5 To 7
    vaSpread1.Row = x
    vaSpread1.Col = 3
    If vaSpread1.Text = "1" Then
      ThisPen = ThisPen + 1
      Thisx = x
    End If
  Next x
  
  If ThisPen > 1 Then
    Call TaxMsg(600, "Only one optional revenue can be earmarked to be used as the penalty revenue. Please review the optional revenues selected to be used for penalty revenues and allow only one selection.")
    Close
    vaTabPro1.ActiveTab = 1
    vaSpread1.SetActiveCell 3, Thisx
    Exit Sub
  End If
  
  If fpcmbRealPersSplitYN.Enabled = False Then
    If Mid(fpcmbRealPersSplitYN.Text, 1, 1) = "Y" Then
      If TaxMsgWOpts(700, "The 'Use Real/Personal Split Billing Y/N?' field is set to 'Yes' but this feature is disabled. If you continue then this setting will be automatically changed to 'No'. Press F10 to continue or press ESC to stop the save procedure safely.", "F10 Continue", "ESC Stop Save") = "abort" Then
        Unload frmTaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        fpcmbCentDepYN.SetFocus
        fpcmbRealPersSplitYN.Enabled = False
        Exit Sub
      Else
        Unload frmTaxMsgWOpts
        fpcmbRealPersSplitYN.Text = "No"
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
        Unload frmTaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        If fptxtCentCash.Enabled = True Then
          fptxtCentCash.SetFocus
        Else
          fpcmbCentDepYN.SetFocus
        End If
        Exit Sub
      Else
        Unload frmTaxMsgWOpts
      End If
    End If
    If QPTrim$(fptxtCentSub.Text) <> "" Then
      If TaxMsgWOpts(700, "You have elected against using Central Depository. However, the Central Depository sub G/L field contains a value. This value will be deleted if you continue to save. Press F10 to continue saving. Otherwise, press ESC to abort save.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmTaxMsgWOpts
        Close
        vaTabPro1.ActiveTab = 1
        If fptxtCentSub.Enabled = True Then
          fptxtCentSub.SetFocus
        Else
          fpcmbCentDepYN.SetFocus
        End If
        Exit Sub
      Else
        Unload frmTaxMsgWOpts
      End If
    End If
  End If
  
  If Exist(TaxSetupName) Then
    EditFlag = True
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    GoSub Check4Penalty
    If TaxMasterRec.LateForm <> CInt(Mid(fpcmbLateFormat.Text, 1, 1)) Then
      If Exist("TXLLPRN.DAT") Then
        If TaxMsgWOpts(700, "There is a late notice letter file saved. It is recommended that this file be deleted before changing to a different late notice letter format. To delete this file press F10. Otherwise press ESC to continue without deleting.", "F10 Delete", "ESC Don't Delete") = "abort" Then
          Unload frmTaxMsgWOpts
          MainLog ("User warned that before changing the late notice format that they allow the program to delete the existing late letter printing records (TXLLPRN.DAT). The user elected NOT to delete.")
        Else
          Unload frmTaxMsgWOpts
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
  TaxMasterRec.CurrYrInt = fptxtCurrYrIntRate.Value
  If Not IsNumeric(fptxtCurrYear.Text) Then
    TaxMasterRec.TaxYear = 0
  Else
    TaxMasterRec.TaxYear = CInt(fptxtCurrYear.Text)
  End If
  
  TaxMasterRec.PastYrInt = fptxtPastYearIntRate.Value
  TaxMasterRec.PenPct = 0 'fptxtPenaltyRate.Value
  
  Select Case fpcmbTaxBillFormat.Text
    Case "MULTI-PART"
      TaxMasterRec.TaxForm = 21837
    Case "POSTCARD"
      TaxMasterRec.TaxForm = 20304
    Case "LASER"
      TaxMasterRec.TaxForm = 16716
    Case "EXPORT REAL"
      TaxMasterRec.TaxForm = 20000
    Case "EXPORT PERSONAL"
      TaxMasterRec.TaxForm = 20001
    Case "HMLT24TF"
      TaxMasterRec.TaxForm = 20002
    Case "PH24TF"
      TaxMasterRec.TaxForm = 20003
    Case "SYL23TF"
      TaxMasterRec.TaxForm = 20004
    Case "BSC32TF"
      TaxMasterRec.TaxForm = 20005
    Case "LLN21TF"
      TaxMasterRec.TaxForm = 20006
    Case "LASER LEGAL"
      TaxMasterRec.TaxForm = 20007
    Case "LASER LEGAL HP"
      TaxMasterRec.TaxForm = 20008
    Case Else
      TaxMasterRec.TaxForm = 0
  End Select
  
  TaxMasterRec.LateForm = CInt(Mid(fpcmbLateFormat.Text, 1, 1))
  TaxMasterRec.DisPct = fptxtDiscPct.Value
  
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
  TaxMasterRec.AutoFillSrvAdd = cmbAutoFill.Text '7/25/08
  
  For x = 1 To 7
    vaSpread1.Row = x
    vaSpread1.Col = 1
    If x = 5 Then
      If QPTrim$(vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev1) = "" Then
          frmTaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(vaSpread1.Text) + " as an optional revenue on row " + CStr(vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmTaxMsgWOpts.Label1.Top = 700
          frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmTaxMsgWOpts.Show vbModal
          If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmTaxMsgWOpts
            Close
            vaTabPro1.ActiveTab = 1
            vaSpread1.SetFocus
            vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        vaSpread1.Col = 3
        If QPTrim$(vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #1 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          vaTabPro1.ActiveTab = 1
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev1) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev1) <> QPTrim$(vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #1. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Close
             vaTabPro1.ActiveTab = 1
             vaSpread1.SetFocus
             vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.OptRev1) Then
                     TblRec.Desc = QPTrim$(vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #1 from " + QPTrim$(TaxMasterRec.OptRev1) + " to " + QPTrim$(vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev1 = QPTrim$(vaSpread1.Text)
    End If
    If x = 6 Then
      If QPTrim$(vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev2) = "" Then
          frmTaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(vaSpread1.Text) + " as an optional revenue on row " + CStr(vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmTaxMsgWOpts.Label1.Top = 700
          frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmTaxMsgWOpts.Show vbModal
          If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmTaxMsgWOpts
            Close
            vaTabPro1.ActiveTab = 1
            vaSpread1.SetFocus
            vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        vaSpread1.Col = 3
        If QPTrim$(vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #2 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          vaTabPro1.ActiveTab = 1
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev2) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev2) <> QPTrim$(vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #2. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmTaxMsgWOpts
             Close
             vaTabPro1.ActiveTab = 1
             vaSpread1.SetFocus
             vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.OptRev2) Then
                     TblRec.Desc = QPTrim$(vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #2 from " + QPTrim$(TaxMasterRec.OptRev2) + " to " + QPTrim$(vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev2 = QPTrim$(vaSpread1.Text)
    End If
    If x = 7 Then
      If QPTrim$(vaSpread1.Text) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev3) = "" Then
          frmTaxMsgWOpts.Label1.Caption = "WARNING: You have elected to save " + QPTrim$(vaSpread1.Text) + " as an optional revenue on row " + CStr(vaSpread1.Row) + ". PLEASE BE ADVISED THAT THIS IS NOT REVERSIBLE. To continue to save press F10. Otherwise press ESC to review."
          frmTaxMsgWOpts.Label1.Top = 700
          frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
          frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
          frmTaxMsgWOpts.Show vbModal
          If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
            Unload frmTaxMsgWOpts
            Close
            vaTabPro1.ActiveTab = 1
            vaSpread1.SetFocus
            vaSpread1.SetActiveCell 1, x
            Exit Sub
          End If
        End If
      Else
        vaSpread1.Col = 3
        If QPTrim$(vaSpread1.Text) = "1" Then
          Call TaxMsg(800, "ERROR: You have selected optional revenue #3 as the penalty revenue but there is no description for this revenue. Please enter a description before continuing.")
          Unload frmTaxMsgWOpts
          vaTabPro1.ActiveTab = 1
          vaSpread1.SetFocus
          vaSpread1.SetActiveCell 1, x
          Exit Sub
        End If
        vaSpread1.Col = 1
      End If
      If EditFlag = True And QPTrim$(TaxMasterRec.OptRev3) <> "" Then
        If QPTrim$(TaxMasterRec.OptRev3) <> QPTrim$(vaSpread1.Text) Then
           If TaxMsgWOpts(600, "You are changing the name of optional revenue #3. All records associated with this revenue will now be reported under the new name. If you wish to continue saving then press F10. Otherwise, press ESC to review and edit the revenue name.", "F10 Continue", "ESC Review") = "abort" Then
             Unload frmTaxMsgWOpts
             Close
             vaTabPro1.ActiveTab = 1
             vaSpread1.SetFocus
             vaSpread1.SetActiveCell 1, x
             Exit Sub
           Else
             OpenTaxRateTables TRHandle, NumOfTRRecs
             For y = 1 To NumOfTRRecs
               Get TRHandle, y, TblRec
                 If TblRec.Deleted <> True Then
                   If QPTrim$(TblRec.Desc) = QPTrim$(TaxMasterRec.OptRev3) Then
                     TblRec.Desc = QPTrim$(vaSpread1.Text)
                     Put TRHandle, y, TblRec
                     MainLog ("User warned about consequences of changing the name of optional revenue #3 from " + QPTrim$(TaxMasterRec.OptRev3) + " to " + QPTrim$(vaSpread1.Text) + " and they saved the new name anyway.")
                   End If
                 End If
              Next y
              Close TRHandle
           End If
        End If
      End If
      TaxMasterRec.OptRev3 = QPTrim$(vaSpread1.Text)
    End If
    
    vaSpread1.Col = 2
    Select Case x
      Case 1
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntIntYN = "Y"
        Else
          TaxMasterRec.IntIntYN = "N"
        End If
      Case 2
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntAdvYN = "Y"
        Else
          TaxMasterRec.IntAdvYN = "N"
        End If
      Case 3
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntLateLstYN = "Y"
        Else
          TaxMasterRec.IntLateLstYN = "N"
        End If
      Case 4
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntPrncTaxYN = "Y"
        Else
          TaxMasterRec.IntPrncTaxYN = "N"
        End If
      Case 5
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt1YN = "Y"
        Else
          TaxMasterRec.IntOpt1YN = "N"
        End If
      Case 6
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt2YN = "Y"
        Else
          TaxMasterRec.IntOpt2YN = "N"
        End If
      Case 7
        If vaSpread1.Text = "1" Then
          TaxMasterRec.IntOpt3YN = "Y"
        Else
          TaxMasterRec.IntOpt3YN = "N"
        End If
    End Select
'    vaSpread1.Col = 3
'    Select Case x
'      Case 5
'        If vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            vaSpread1.Col = 1
'            frmTaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmTaxMsgWOpts.Label1.Top = 700
'            frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmTaxMsgWOpts.Show vbModal
'            If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmTaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmTaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 5
'        End If
'      Case 6
'        If vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            frmTaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmTaxMsgWOpts.Label1.Top = 700
'            frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmTaxMsgWOpts.Show vbModal
'            If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmTaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmTaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 6
'        End If
'      Case 7
'        If vaSpread1.Text = "1" Then
'          If CDbl(fptxtPenaltyRate.Text) = 0 Then
'            frmTaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is set to zero. If you wish to save anyway then press F10. Otherwise press ESC to review."
'            frmTaxMsgWOpts.Label1.Top = 700
'            frmTaxMsgWOpts.cmdCont.Text = "F10 Save"
'            frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
'            frmTaxMsgWOpts.Show vbModal
'            If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
'              Unload frmTaxMsgWOpts
'              Close
'              vaTabPro1.ActiveTab = 1
'              fptxtPenaltyRate.SetFocus
'              Exit Sub
'            Else
'              Unload frmTaxMsgWOpts
'            End If
'          End If
'          TaxMasterRec.PenIdx = 7
'        End If
'    End Select
    vaSpread1.Col = 4
    Select Case x
      Case 1
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenIntYN = "Y"
        Else
          TaxMasterRec.PenIntYN = "N"
        End If
      Case 2
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenAdvYN = "Y"
        Else
          TaxMasterRec.PenAdvYN = "N"
        End If
      Case 3
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenLateLstYN = "Y"
        Else
          TaxMasterRec.PenLateLstYN = "N"
        End If
      Case 4
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenPrncTaxYN = "Y"
        Else
          TaxMasterRec.PenPrncTaxYN = "N"
        End If
      Case 5
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt1YN = "Y"
        Else
          TaxMasterRec.PenOpt1YN = "N"
        End If
      Case 6
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt2YN = "Y"
        Else
          TaxMasterRec.PenOpt2YN = "N"
        End If
      Case 7
        If vaSpread1.Text = "1" Then
          TaxMasterRec.PenOpt3YN = "Y"
        Else
          TaxMasterRec.PenOpt3YN = "N"
        End If
    End Select
  Next x
  If EditFlag = False Then
    TaxMasterRec.DiscXDate = 0
  End If
  TaxMasterRec.UseCyclesYN = Mid(fpcmbCyclesYN.Text, 1, 1)
  TaxMasterRec.UseCountyYN = Mid(fpcmbCountyYN.Text, 1, 1)
  TaxMasterRec.RealPersSplit = Mid(fpcmbRealPersSplitYN.Text, 1, 1)
  Put TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Call LogSaves
  
  Unload frmTaxGLList
  Call Savemsg(900, "Your data has been saved successfully.")
  
  If Exist("C:\CPWork\ratetbls.dat") Then
    KillFile "C:\CPWork\ratetbls.dat"
    vaSpread1.Col = 1
    For x = 5 To 7
      vaSpread1.Row = x
      If QPTrim$(vaSpread1.Text) <> "" Then
        frmTaxRateMenu.Show
        DoEvents
        Unload Me
        Exit Sub
      End If
    Next x
  End If
  
  If Exist("C:\CPWork\lateltr.dat") Then
    KillFile "C:\CPWork\lateltr.dat"
    frmTaxBillingMenu.Show
    DoEvents
    Unload Me
  Else
    frmTaxBillSetUpMenu.Show
    DoEvents
    Unload Me
  End If
  
  Exit Sub
  
Check4Penalty:
  PenCnt = 0
  If PenIdx = 0 Then Return
  If TaxMasterRec.PenIdx = 0 Then Return
  
  vaSpread1.Col = 2
  For y = 5 To 7
    vaSpread1.Row = y
    If vaSpread1.Text = "1" Then
      PenCnt = PenCnt + 1
        vaSpread1.Col = 1
        ThisDesc$ = QPTrim$(vaSpread1.Text)
        If ThisDesc = "" Then
          frmTaxMsg.Label1.Caption = "You have designated revenue #" + CStr(y) + " as your penalty revenue but no description has been entered. Please enter a description for this penalty revenue."
          frmTaxMsg.Show vbModal
          Close
          vaSpread1.SetActiveCell 1, y
          Exit Sub
        End If
    End If
  Next y
    
'  If PenCnt > 0 And fptxtPenaltyRate.Value = 0 Then
'    frmTaxMsgWOpts.Label1.Caption = "You have elected to assess penalties but the penalty rate is zero. Press F10 if you wish to continue saving anyway. Press ESC to review."
'    frmTaxMsgWOpts.Label1.Top = 800
'    frmTaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
'    frmTaxMsgWOpts.cmdExit.Text = "ESC Review"
'    frmTaxMsgWOpts.Show vbModal
'    If frmTaxMsgWOpts.fptxtChoice.Text = "abort" Then
'      Unload frmTaxMsgWOpts
'      fptxtPenaltyRate.SetFocus
'      Close
'      Exit Sub
'    ElseIf frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
'      Unload frmTaxMsgWOpts
'    End If
'  End If
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "cmdSave", Erl)
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
    frmTempTaxBillPrint.Show
  ElseIf fpcmbTaxBillFormat.Text = "POSTCARD" Then
    frmTaxPostCard1.Show
  ElseIf fpcmbTaxBillFormat.Text = "HMLT24TF" Then 'Hamlet 24 inch tractor fed
    frmTaxBillHamlet.Show
  ElseIf fpcmbTaxBillFormat.Text = "PH24TF" Then 'Pink Hill 24 inch tractor fed
    frmTaxBillPH24TF.Show
  ElseIf fpcmbTaxBillFormat.Text = "SYL23TF" Then 'Sylva 23 inch tractor fed
    frmTaxBillSYL23TF.Show
  ElseIf fpcmbTaxBillFormat.Text = "BSC32TF" Then 'Biscoe 32 inch tractor fed
    frmTaxBillBSC32TF.Show
  ElseIf fpcmbTaxBillFormat.Text = "LLN21TF" Then 'Leland 21 inch tractor fed
    frmTaxBillLLN21TF.Show
  ElseIf fpcmbTaxBillFormat.Text = "LASER LEGAL" Then 'Laser legal size
    Call ShowLaserLegal
  ElseIf fpcmbTaxBillFormat.Text = "LASER LEGAL HP" Then 'Laser legal size HP Printer
    Call ShowLaserLegal
  ElseIf InStr(fpcmbTaxBillFormat.Text, "EXPORT") Then
    Call TaxMsg(800, "The 'EXPORT' bill format creates a file that is then forwarded to a tax billing company. There is no hard copy tax bill to display.")
    Exit Sub
  End If
  
  DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmTaxSystemSetup.")
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
    Case vbKeyF3:
      SendKeys "%S"
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
      Unload frmTaxGLList
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxSystemSetup.")
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
  
  'on error goto ERRORSTUFF
  
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
    Case 21837
      fpcmbTaxBillFormat.Text = "MULTI-PART"
    Case 20304
      fpcmbTaxBillFormat.Text = "POSTCARD"
    Case 16716
      fpcmbTaxBillFormat.Text = "LASER"
    Case 20000
      fpcmbTaxBillFormat.Text = "EXPORT REAL"
    Case 20001
      fpcmbTaxBillFormat.Text = "EXPORT PERSONAL"
    Case 20002
      fpcmbTaxBillFormat.Text = "HMLT24TF"
    Case 20003
      fpcmbTaxBillFormat.Text = "PH24TF"
    Case 20004
      fpcmbTaxBillFormat.Text = "SYL23TF"
    Case 20005
      fpcmbTaxBillFormat.Text = "BSC32TF"
    Case 20006
      fpcmbTaxBillFormat.Text = "LLN21TF"
    Case 20007
      fpcmbTaxBillFormat.Text = "LASER LEGAL"
    Case 20008
      fpcmbTaxBillFormat.Text = "LASER LEGAL HP"
    Case Else
      fpcmbTaxBillFormat.Text = "UNKNOWN"
  End Select
  TempTaxForm = TaxMasterRec.TaxForm
  
  fpcmbTaxBillFormat.AddItem "POSTCARD"
  fpcmbTaxBillFormat.AddItem "LASER"
  fpcmbTaxBillFormat.AddItem "MULTI-PART"
  fpcmbTaxBillFormat.AddItem "HMLT24TF"
  fpcmbTaxBillFormat.AddItem "PH24TF"
  fpcmbTaxBillFormat.AddItem "SYL23TF"
  fpcmbTaxBillFormat.AddItem "BSC32TF"
  fpcmbTaxBillFormat.AddItem "LLN21TF"
  fpcmbTaxBillFormat.AddItem "LASER LEGAL"
  fpcmbTaxBillFormat.AddItem "LASER LEGAL HP"
  fpcmbTaxBillFormat.AddItem "EXPORT REAL"
  fpcmbTaxBillFormat.AddItem "EXPORT PERSONAL"
  
  fpcmbStateOfTax.AddItem "NC"
  fpcmbStateOfTax.AddItem "VA"
  fpcmbStateOfTax.AddItem "GA"
  fpcmbStateOfTax.AddItem "SC"
  
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
  
  fptxtCurrYrIntRate.Text = TaxMasterRec.CurrYrInt
  TempCurrYrInt = TaxMasterRec.CurrYrInt
  
  DateLen = Len(Date)
  ThisYr = Mid(Date, DateLen - 3, DateLen)
  If TaxMasterRec.TaxYear = 0 Then
    fptxtCurrYear.Text = ThisYr
    TempTaxYear = 0
  Else
    fptxtCurrYear.Text = CStr(TaxMasterRec.TaxYear)
    TempTaxYear = TaxMasterRec.TaxYear
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
  
  fptxtDiscPct.Text = TaxMasterRec.DisPct
  TempDisPct = TaxMasterRec.DisPct
  
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
  
  If TaxMasterRec.AutoFillSrvAdd = "N" Then '7/25/08
    cmbAutoFill.Text = "No"
  Else
    cmbAutoFill.Text = "Yes"
  End If
  
  cmbAutoFill.AddItem "Yes" '7/25/08
  cmbAutoFill.AddItem "No" '7/25/08
  
  TempWarnInt = fpcmbNoInterYN.Text
  
  fptxtOverPayGL.Text = QPTrim$(TaxMasterRec.OverPayGLNum)
  TempOverPayGLNum = QPTrim$(TaxMasterRec.OverPayGLNum)
  
  For x = 1 To 7
    vaSpread1.Row = x
    vaSpread1.Col = 1
    Select Case x
      Case 1
        vaSpread1.Text = "Default: Interest Accrued"
        vaSpread1.Lock = True
      Case 2
        vaSpread1.Text = "Default: Advertising Cost Incurred"
        vaSpread1.Lock = True
      Case 3
        vaSpread1.Text = "Default: Late Listing"
        vaSpread1.Lock = True
      Case 4
        vaSpread1.Text = "Default: Principal"
        vaSpread1.Lock = True
      Case 5
        vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev1)
        TempOptRev1 = TaxMasterRec.OptRev1
      Case 6
        vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev2)
        TempOptRev2 = TaxMasterRec.OptRev2
      Case 7
        vaSpread1.Text = QPTrim$(TaxMasterRec.OptRev3)
        TempOptRev3 = TaxMasterRec.OptRev3
      Case Else
    End Select
    vaSpread1.Col = 2
    Select Case x
      Case 1
        If TaxMasterRec.IntIntYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntIntYN = TaxMasterRec.IntIntYN
      Case 2
        If TaxMasterRec.IntAdvYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntAdvYN = TaxMasterRec.IntAdvYN
      Case 3
        If TaxMasterRec.IntLateLstYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntLateLstYN = TaxMasterRec.IntLateLstYN
      Case 4
        If TaxMasterRec.IntPrncTaxYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntPrncTaxYN = TaxMasterRec.IntPrncTaxYN
      Case 5
        If TaxMasterRec.IntOpt1YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntOpt1YN = TaxMasterRec.IntOpt1YN
      Case 6
        If TaxMasterRec.IntOpt2YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntOpt2YN = TaxMasterRec.IntOpt2YN
      Case 7
        If TaxMasterRec.IntOpt3YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempIntOpt3YN = TaxMasterRec.IntOpt3YN
      Case Else
    End Select
    
    vaSpread1.Col = 3
    Select Case x
      Case 1
        vaSpread1.Lock = True
      Case 2
        vaSpread1.Lock = True
      Case 3
        vaSpread1.Lock = True
      Case 4
        vaSpread1.Lock = True
      Case 5
        If PenIdx = 5 Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        
      Case 6
        If TaxMasterRec.PenIdx = 6 Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
      Case 7
        If TaxMasterRec.PenIdx = 7 Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
      Case Else
    End Select
    
    vaSpread1.Col = 4
    Select Case x
      Case 1
        If TaxMasterRec.PenIntYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenIntYN = TaxMasterRec.PenIntYN
      Case 2
        If TaxMasterRec.PenAdvYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenAdvYN = TaxMasterRec.PenAdvYN
      Case 3
        If TaxMasterRec.PenLateLstYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenLateLstYN = TaxMasterRec.PenLateLstYN
      Case 4
        If TaxMasterRec.PenPrncTaxYN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenPrncTaxYN = TaxMasterRec.PenPrncTaxYN
      Case 5
        If TaxMasterRec.PenOpt1YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenOpt1YN = TaxMasterRec.PenOpt1YN
      Case 6
        If TaxMasterRec.PenOpt2YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenOpt2YN = TaxMasterRec.PenOpt2YN
      Case 7
        If TaxMasterRec.PenOpt3YN = "Y" Then
          vaSpread1.Value = 1
        Else
          vaSpread1.Value = 0
        End If
        TempPenOpt3YN = TaxMasterRec.PenOpt3YN
      Case Else
    End Select
  Next x
  
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
  
  fpcmbRealPersSplitYN.AddItem "Yes"
  fpcmbRealPersSplitYN.AddItem "No"
  If TaxMasterRec.RealPersSplit <> "Y" Then
    fpcmbRealPersSplitYN.Text = "No"
  Else
    fpcmbRealPersSplitYN.Text = "Yes"
  End If
  TempRealPersSplit = TaxMasterRec.RealPersSplit
  
  Call FixSpread
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "LoadMe", Erl)
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
        fptxtDiscPct.SetFocus
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
      If fpcmbRealPersSplitYN.Enabled = True Then
        fpcmbRealPersSplitYN.SetFocus
      Else
        fptxtDiscPct.SetFocus
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


Private Sub fpcmbRealPersSplitYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRealPersSplitYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRealPersSplitYN.ListIndex = -1
  End If
  If fpcmbRealPersSplitYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtDiscPct.SetFocus
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
    fpcmbRealPersSplitYN.Enabled = False
  Else
    fpcmbRealPersSplitYN.Enabled = True
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
      fptxtCurrYrIntRate.SetFocus
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
  
  'on error goto ERRORSTUFF
  
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Name of Taxing Authority' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Address 1' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Address 2' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'City' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'State' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Zip Code' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'State of Tax' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Customer Related Description' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchCust = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Customer Related Description has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPropOptSrch
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.OptSrchProp)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Property Related Description' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchProp = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Property Related Description has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPersOptSrch
  TabNum = 0
  ThisDesc = QPTrim$(TaxRec.OptSrchPers)
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Personal Related Description' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.OptSrchPers = QPTrim$(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Personal Related Description has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrYrIntRate
  TabNum = 0
  ThisDbl = TaxRec.CurrYrInt
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Current Year Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.CurrYrInt = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Current Year Interest Rate has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtCurrYear
  TabNum = 0
  OptInt = TaxRec.TaxYear
  If CInt(ThisControl.Text) <> OptInt Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Current Tax Year' field has been changed from " + CStr(OptInt) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.TaxYear = CInt(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Current Tax Year has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtPastYearIntRate
  TabNum = 0
  ThisDbl = TaxRec.PastYrInt
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Past Year Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Minimum Tax Amount' field has been changed from " + Using("$##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Minimum Tax Options' field has been changed from " + ThisDesc + " to " + OptStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
'    frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Interest Rate' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmTaxMsgW4Opts.Label1.Top = 575
'    frmTaxMsgW4Opts.Show vbModal
'    choice = frmTaxMsgW4Opts.fptxtChoice.Text
'    Unload frmTaxMsgW4Opts
'    If choice = "save" Then
'      TaxRec.PenPct = CDbl(ThisControl.Text)
'      Put TMHandle, 1, TaxRec
'      Call Savemsg(900, "Penalty Interest Rate has been saved successfully.")
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fpcmbCentDepYN
  TabNum = 1
  ThisDesc = TaxRec.CntrlDepYN
  If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
    If ThisDesc = "N" Then
      ThisDesc = "No"
    ElseIf ThisDesc = "Y" Then
      ThisDesc = "Yes"
    End If
    frmTaxMsgW4Opts.Label1.Caption = "The 'Central Depository Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Central Depository Cash G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtCentCash.Text)) = False Then
        frmTaxMsgWOpts.Label1.Caption = "The Central Depository Cash G/L number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmTaxMsgWOpts.Label1.Top = 600
        frmTaxMsgWOpts.Show vbModal
        If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmTaxMsgWOpts
          MainLog ("Warning: User issued warning that the central depository cash GL number " + QPTrim$(fptxtCentCash.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmTaxMsgWOpts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Central Depository Sub G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtCentSub.Text)) = False Then
        frmTaxMsgWOpts.Label1.Caption = "The Central Depository Sub G/L number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmTaxMsgWOpts.Label1.Top = 600
        frmTaxMsgWOpts.Show vbModal
        If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmTaxMsgWOpts
          MainLog ("Warning: User issued warning that the central depository sub GL number " + QPTrim$(fptxtCentSub.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmTaxMsgWOpts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'No Interest Warning Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.WarnInt = ThisControl.Text
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "No Interest Warning Y/N? has been saved successfully.")
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Accounting Method' field has been changed from " + ThisDesc + " to " + OptStr + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.AcctgMethod = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Accounting Method has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
 
  Select Case TaxRec.TaxForm
    Case 21837
      ThisDesc = "MULTI-PART"
    Case 20304
      ThisDesc = "POSTCARD"
    Case 16716
      ThisDesc = "LASER"
    Case 20000
      ThisDesc = "EXPORT REAL"
    Case 20001
      ThisDesc = "EXPORT PERSONAL"
    Case 20002
      ThisDesc = "HMLT24TF"
    Case 20003
      ThisDesc = "PH24TF"
    Case 20004
      ThisDesc = "SYL23TF"
    Case 20005
      ThisDesc = "BSC32TF"
    Case 20006
      ThisDesc = "LLN21TF"
    Case 20007
      ThisDesc = "LASER LEGAL"
    Case 20008
      ThisDesc = "LASER LEGAL HP"
    Case Else
      ThisDesc = "UNKNOWN"
  End Select
  
  Select Case QPTrim$(fpcmbTaxBillFormat.Text)
    Case "MULTI-PART"
      OptInt = 21837
    Case "POSTCARD"
      OptInt = 20304
    Case "LASER"
      OptInt = 16716
    Case "HMLT24TF"
      OptInt = 20002
    Case "PH24TF"
      OptInt = 20003
    Case "SYL23TF"
      OptInt = 20004
    Case "BSC32TF"
      OptInt = 20005
    Case "LLN21TF"
      OptInt = 20006
    Case "LASER LEGAL"
      OptInt = 20007
    Case "LASER LEGAL HP"
      OptInt = 20008
    Case "EXPORT REAL"
      OptInt = 20000
    Case "EXPORT PERSONAL"
      OptInt = 20001
    Case Else
      OptInt = 0
  End Select
  
  Set ThisControl = fpcmbTaxBillFormat
  TabNum = 1
  If QPTrim$(ThisControl.Text) <> ThisDesc Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Tax Bill Format' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Late Bill Format' field has been changed from " + CStr(OptInt) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Overpayment G/L Number' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      If VerifyGLNum(QPTrim$(fptxtOverPayGL.Text)) = False Then
        frmTaxMsgWOpts.Label1.Caption = "The Overpayment GL number could not be located in the current GL index file. If you wish to save it anyway then press F10. Otherwise, press ESC to return to the screen without saving."
        frmTaxMsgWOpts.Label1.Top = 600
        frmTaxMsgWOpts.Show vbModal
        If frmTaxMsgWOpts.fptxtChoice.Text = "continue" Then
          Unload frmTaxMsgWOpts
          MainLog ("Warning: User issued warning that the overpayment GL number " + QPTrim$(fptxtOverPayGL.Text) + " could not be verified and they elected to continue to save it anyway.")
        Else
          Unload frmTaxMsgWOpts
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
'    frmTaxMsgW4Opts.Label1.Caption = "The 'Do you use multiple revenue accounts for prior years Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmTaxMsgW4Opts.Label1.Top = 575
'    frmTaxMsgW4Opts.Show vbModal
'    choice = frmTaxMsgW4Opts.fptxtChoice.Text
'    Unload frmTaxMsgW4Opts
'    If choice = "save" Then
'      TaxRec.PriorYrMltRevYN = ThisControl.Text
'      Put TMHandle, 1, TaxRec
'      Call Savemsg(900, "Do you use multiple revenue accounts for prior years Y/N? has been saved successfully.")
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fptxtDiscPct
  TabNum = 1
  ThisDbl = TaxRec.DisPct
  If CDbl(ThisControl.Text) <> ThisDbl Then
    frmTaxMsgW4Opts.Label1.Caption = "The 'Discount Percentage' field has been changed from " + Using("##0.00", ThisDbl) + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.DisPct = CDbl(ThisControl.Text)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Discount Percentage has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Dim SpreadText As String
  Dim ThisCol As Integer
  
  ThisCol = 1
  vaSpread1.Col = ThisCol
  For x = 5 To 7
    vaSpread1.Row = x
    If x = 5 Then
      SpreadText = QPTrim$(vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev1)
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #1' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmTaxMsgW4Opts.Label1.Top = 375
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
    If x = 6 Then
      SpreadText = QPTrim$(vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev2)
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #2' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmTaxMsgW4Opts.Label1.Top = 375
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
    If x = 7 Then
      SpreadText = QPTrim$(vaSpread1.Text)
      ThisDesc = QPTrim$(TaxRec.OptRev3)
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Optional Revenue #3' field has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes. SAVING ASSOCIATES ALL CURRENT RECORDS WITH THE NEW NAME."
        frmTaxMsgW4Opts.Label1.Top = 375
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
  vaSpread1.Col = ThisCol
  For x = 1 To 7
    vaSpread1.Row = x
    If vaSpread1.Text = "1" Then
      SpreadText = "Y"
    Else
      SpreadText = "N"
    End If
    vaSpread1.Col = 1
    SpreadText2 = QPTrim$(vaSpread1.Text)
    vaSpread1.Col = ThisCol
    If x = 1 Then
      ThisDesc = TaxRec.IntIntYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
      ThisDesc = TaxRec.IntPrncTaxYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntPrncTaxYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 5 Then
      ThisDesc = TaxRec.IntOpt1YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntOpt1YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 6 Then
      ThisDesc = TaxRec.IntOpt2YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.IntOpt2YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Apply Interest field for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 7 Then
      ThisDesc = TaxRec.IntOpt3YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Apply Interest' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
  vaSpread1.Col = ThisCol
  For x = 5 To 7
    vaSpread1.Row = x
    If vaSpread1.Text = "1" Then
      SpreadText = "Y"
      If TaxRec.PenIdx <> x Then
        ThisDesc = "N"
        vaSpread1.Col = 1
        SpreadText2 = QPTrim$(vaSpread1.Text)
        vaSpread1.Col = ThisCol
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalty Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
  vaSpread1.Col = ThisCol
  For x = 1 To 7
    vaSpread1.Row = x
    If vaSpread1.Text = "1" Then
      SpreadText = "Y"
    Else
      SpreadText = "N"
    End If
    vaSpread1.Col = 1
    SpreadText2 = QPTrim$(vaSpread1.Text)
    vaSpread1.Col = ThisCol
    If x = 1 Then
      ThisDesc = TaxRec.PenIntYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
      ThisDesc = TaxRec.PenPrncTaxYN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenPrncTaxYN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 5 Then
      ThisDesc = TaxRec.PenOpt1YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenOpt1YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 6 Then
      ThisDesc = TaxRec.PenOpt2YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
        If choice = "save" Then
          TaxRec.PenOpt2YN = SpreadText
          Put TMHandle, 1, TaxRec
          Call Savemsg(900, "Penalize This Rev for " + SpreadText2 + " has been saved successfully.")
        Else
          GoSub HandleChoice
        End If
      End If
    End If
    If x = 7 Then
      ThisDesc = TaxRec.PenOpt3YN
      If ThisDesc = "N" Then
        ThisDesc = "N"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Y"
      End If
      If SpreadText <> ThisDesc Then
        frmTaxMsgW4Opts.Label1.Caption = "The 'Penalize This Rev' field for " + SpreadText2 + " has been changed from " + ThisDesc + " to " + SpreadText + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
        frmTaxMsgW4Opts.Label1.Top = 575
        frmTaxMsgW4Opts.Show vbModal
        choice = frmTaxMsgW4Opts.fptxtChoice.Text
        Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Use Billing Cycles Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
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
    frmTaxMsgW4Opts.Label1.Caption = "The 'Use County Billing Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmTaxMsgW4Opts.Label1.Top = 575
    frmTaxMsgW4Opts.Show vbModal
    choice = frmTaxMsgW4Opts.fptxtChoice.Text
    Unload frmTaxMsgW4Opts
    If choice = "save" Then
      TaxRec.UseCountyYN = Mid(ThisControl.Text, 1, 1)
      Put TMHandle, 1, TaxRec
      Call Savemsg(900, "Use County Billing Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
    
  If fpcmbRealPersSplitYN.Enabled = True Then
    Set ThisControl = fpcmbRealPersSplitYN
    TabNum = 1
    ThisDesc = TaxRec.RealPersSplit
    If Mid(ThisControl.Text, 1, 1) <> ThisDesc Then
      If ThisDesc = "N" Then
        ThisDesc = "No"
      ElseIf ThisDesc = "Y" Then
        ThisDesc = "Yes"
      End If
      frmTaxMsgW4Opts.Label1.Caption = "The 'Use Real/Personal Split Billing Y/N?' field has been changed from " + ThisDesc + " to " + QPTrim$(ThisControl.Text) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmTaxMsgW4Opts.Label1.Top = 575
      frmTaxMsgW4Opts.Show vbModal
      choice = frmTaxMsgW4Opts.fptxtChoice.Text
      Unload frmTaxMsgW4Opts
      If choice = "save" Then
        TaxRec.RealPersSplit = Mid(ThisControl.Text, 1, 1)
        Put TMHandle, 1, TaxRec
        Call Savemsg(900, "Use Real/Personal Split Billing Y/N? has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Close TMHandle
  
  Exit Function
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close TMHandle
        frmTaxBillSetUpMenu.Show
        DoEvents
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        vaTabPro1.ActiveTab = TabNum
        If x > 0 Then
          vaSpread1.SetActiveCell ThisCol, x
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "Check4Changes", Erl)
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
    MainLog ("frmTaxSystemSetup: Name of Taxing Authority was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempADD1)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Add1)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Address #1 was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempADD2)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Add2)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Address #2 was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempCity)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.City)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: City was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempTownState)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.TownState)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The town's state was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempZip)
  ThisZip = ReplaceString(TempStr, "-", "")
  If QPTrim$(ThisZip) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.Zip)
  ThatZip = ReplaceString(TempSave, "-", "")
  If QPTrim$(ThatZip) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Zip Code was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempTaxSt)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.TaxSt)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The tax state was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchCust)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchCust)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The optional customer search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchProp)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchProp)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The optional property search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = QPTrim$(TempOptSrchPers)
  If TempStr = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptSrchPers)
  If TempSave = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The optional personal search field was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempDbl = TempCurrYrInt
  TempSaveDbl = TaxRec.CurrYrInt
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmTaxSystemSetup: The current year's tax percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempInt = TempTaxYear
  TempSaveInt = TaxRec.TaxYear
  If TempInt <> TempSaveInt Then
    MainLog ("frmTaxSystemSetup: The current tax year was changed from " + CStr(TempInt) + " to " + CStr(TempSaveInt) + " and saved.")
  End If
    
  TempDbl = TempPastYrInt
  TempSaveDbl = TaxRec.PastYrInt
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmTaxSystemSetup: The past year's tax percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
    
  TempDbl = TempPenPct
  TempSaveDbl = TaxRec.PenPct
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmTaxSystemSetup: The penalty percentage was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  Select Case TempTaxForm
    Case 20304
      TempStr = "POSTCARD"
    Case 21837
      TempStr = "MULTI-PART"
    Case 16716
      TempStr = "LASER"
    Case 20002
      TempStr = "HMLT24TF"
    Case 20003
      TempStr = "PH24TF"
    Case 20004
      TempStr = "SYL23TF"
    Case 20005
      TempStr = "BSC32TF"
    Case 20006
      TempStr = "LLN21TF"
    Case 20007
      TempStr = "LASER LEGAL"
    Case 20008
      TempStr = "LASER LEGAL HP"
    Case Else
      TempStr = "UNKNOWN"
  End Select
  Select Case TaxRec.TaxForm
    Case 20304
      TempSave = "POSTCARD"
    Case 21837
      TempSave = "MULTI-PART"
    Case 16716
      TempSave = "LASER"
    Case 20002
      TempStr = "HMLT24TF"
    Case 20003
      TempStr = "PH24TF"
    Case 20004
      TempStr = "SYL23TF"
    Case 20005
      TempStr = "BSC32TF"
    Case 20006
      TempStr = "LLN21TF"
    Case 20007
      TempStr = "LASER LEGAL"
    Case 20008
      TempStr = "LASER LEGAL HP"
    Case Else
      TempSave = "UNKNOWN"
  End Select
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: The tax bill format was changed from " + TempStr + " to " + TempSave + " and saved.")
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
    MainLog ("frmTaxSystemSetup: The minimum tax option was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempDbl = TempMinTxPct
  TempSaveDbl = TaxRec.MinBill
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmTaxSystemSetup: The minimum tax amount was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
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
    MainLog ("frmTaxSystemSetup: The accounting method was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  If TempDisPct = "" Then TempDisPct = "0"
  TempDbl = TempDisPct
  TempSaveDbl = TaxRec.DisPct
  If TempDbl <> TempSaveDbl Then
    MainLog ("frmTaxSystemSetup: The discount amount was changed from " + Using("##0.00", TempDbl) + " to " + Using("##0.00", TempSaveDbl) + " and saved.")
  End If
  
  TempStr = TempCntrlDepYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.CntrlDepYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Central Depository Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempCDCashGL)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.CDCashGL)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Central Depository Cash G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempCDSubGL)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.CDSubGL)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Central Depository Sub G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempPriorYrMltRevYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PriorYrMltRevYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Do you use multiple revenue accounts for prior years Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOverPayGLNum)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OverPayGLNum)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Overpayment G/L Number was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempPenPrncTaxYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenPrncTaxYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for Principle was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenIntYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenIntYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for Interest Accrued was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenAdvYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenAdvYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for Advertising was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If

  TempStr = TempPenLateLstYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenLateLstYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for Late Listing was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempWarnInt
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.WarnInt
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: No Interest Warning Y/N was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev1)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev1)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Revenue Description for Optional Revenue #1 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
'  vaSpread1.Col = 1
'  vaSpread1.Row = 5
  TempStr = TempPenOpt1YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt1YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for " + QPTrim$(vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev2)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev2)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Revenue Description for Optional Revenue #2 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
'  vaSpread1.Col = 1
'  vaSpread1.Row = 6
  TempStr = TempPenOpt2YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt2YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for " + QPTrim$(vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = QPTrim$(TempOptRev3)
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = QPTrim$(TaxRec.OptRev3)
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Revenue Description for Optional Revenue #3 was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
    
  TempStr = TempPenOpt3YN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.PenOpt3YN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Penalize This Rev for " + QPTrim$(vaSpread1.Text) + " was changed from  " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  If TaxRec.PenIdx > 0 Then
    vaSpread1.Col = 1
    vaSpread1.Row = TaxRec.PenIdx
    MainLog ("frmTaxSystemSetup: Penalty revenue saved as " + QPTrim$(vaSpread1.Text) + " on row # " + CStr(TaxRec.PenIdx) + ".")
  Else
    MainLog ("frmTaxSystemSetup: No penalty revenue index was saved.")
  End If
  
  TempStr = TempUseCyclesYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.UseCyclesYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Use Billing Cycles Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempStr = TempUseCountyYN
  If QPTrim$(TempStr) = "" Then TempStr = "BLANK"
  TempSave = TaxRec.UseCountyYN
  If QPTrim$(TempSave) = "" Then TempSave = "BLANK"
  If TempStr <> TempSave Then
    MainLog ("frmTaxSystemSetup: Use County Billing Y/N? was changed from " + TempStr + " to " + TempSave + " and saved.")
  End If
  
  TempInt = TempRealPersSplit
  TempSaveInt = TaxRec.RealPersSplit
  If TempInt <> TempSaveInt Then
    MainLog ("frmTaxSystemSetup: Use Real/Pers Split Billing Y/N? was changed from " + CStr(TempInt) + " to " + CStr(TempSaveInt) + " and saved.")
  End If
  
End Sub

Private Sub fpListTownships_DblClick()
  cmdAddTownship.Text = "Edit Township"
  TSListIdx = fpListTownships.ListIndex
  fptxtTownShipName.Text = QPTrim$(fpListTownships.Text)
End Sub

Private Sub fptxtCentCash_DblClick(Button As Integer)
  fptxtCentCash.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtCentSub_DblClick(Button As Integer)
  fptxtCentSub.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtCustOptSrch_Change()
  If InStr(fptxtCustOptSrch.Text, " ORDER") Then
    If TaxMsgWOpts(750, "Using the word 'Order' in the description will be displayed as " + fptxtCustOptSrch.Text + " Order in the Print Order drop down boxes. If this is OK then press F10 to approve. Otherwise press ESC to return and edit.", "F10 Approve", "ESC Edit") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtCustOptSrch.SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub fptxtDiscPct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbCentDepYN.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpcmbRealPersSplitYN.Enabled = True Then
      fpcmbRealPersSplitYN.SetFocus
    Else
      fpcmbCountyYN.SetFocus
    End If
  End If
End Sub

Private Sub fptxtOverPayGL_DblClick(Button As Integer)
  fptxtOverPayGL.Text = Clipboard.GetText
  frmTaxGLList.ZOrder 0
End Sub

Private Sub fptxtOverPayGL_LostFocus()
  Dim ThisLen As Integer
  Dim ThatLen As Integer
  
  If QPTrim$(fptxtOverPayGL.Text) = "" Then Exit Sub
  ThatLen = Fund + Dept + Detail
  If ThatLen = 0 Then Exit Sub
  fptxtOverPayGL.Text = ReplaceString(fptxtOverPayGL.Text, "-", "")
  ThisLen = Len(QPTrim$(fptxtOverPayGL.Text))
  If ThisLen <> ThatLen Then
    If TaxMsgWOpts(750, "The GL number entered is not the same length as the other GL numbers saved. If you wish to review this entry press ESC. Otherwise, press F10 to continue.", "F10 Continue Anyway", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtOverPayGL.SetFocus
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      Exit Sub
    End If
  End If
  fptxtOverPayGL.Text = AddDashesToGLNumber(fptxtOverPayGL.Text, Fund, Dept, Detail)
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
  Dim PenCnt As Integer
  Dim x As Integer
  Dim CntPens As Integer
  Dim Thisx As Integer
  
  'on error goto ERRORSTUFF
  
  StrEmpty = False
  vaSpread1.Col = 1
  vaSpread1.Row = Row
  If QPTrim$(vaSpread1.Text) = "" Then
    StrEmpty = True
  End If
  vaSpread1.Col = 2
  If vaSpread1.Text = "1" And StrEmpty = True Then
    Call TaxMsg(800, "This row contains an unused optional revenue. Setting interest is not allowed for this row.")
    vaSpread1.Text = "0"
    vaTabPro1.ActiveTab = 1
    vaSpread1.SetFocus
    vaSpread1.SetActiveCell 2, Row
    Exit Sub
  End If
  
  If Col <> 3 Then Exit Sub
  CntPens = 0
  Thisx = 0
  vaSpread1.Col = 3
  For x = 5 To 7
    vaSpread1.Row = x
    If vaSpread1.Text = "1" Then
      CntPens = CntPens + 1
      Thisx = x
    End If
  Next x
  
  If CntPens > 1 Then
    Call TaxMsg(800, "ERROR: Only one optional revenue can be earmarked as the penalty revenue. Please review your penalty selections and select only one penalty revenue.")
    vaTabPro1.ActiveTab = 1
    vaSpread1.SetActiveCell 3, Thisx
    Exit Sub
  End If
  
  vaSpread1.Col = Col
  vaSpread1.Row = Row
    
    
  If vaSpread1.Text = "1" And PenIdx <> Row Then
    vaSpread1.Col = 1
    If QPTrim$(vaSpread1.Text) = "" Then
      Call TaxMsg(800, "The penalty revenue source has been assigned to row " + CStr(Row) + ". Please enter a penalty description.")
      vaTabPro1.ActiveTab = 1
      vaSpread1.SetActiveCell 1, Row
    End If
    vaSpread1.Col = Col
    If PenIdx > 0 Then
      vaSpread1.Row = PenIdx
      vaSpread1.Value = "0"
    End If
    PenIdx = Row
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "vaSpread1_Change", Erl)
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

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

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
   
   'on error goto ERRORSTUFF
   
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSystemSetup", "VerifyGLNum", Erl)
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
    frmTaxBillSetUpMenu.Show
    DoEvents
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
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 16
          Next y
        Next x
        vaSpread1.RowHeight(-1) = 27.5
        vaSpread1.RowHeight(0) = 27.5
      Else
        COne = 11.25
        coladj = 3.45
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 14
          Next y
        Next x
        vaSpread1.Col = 2
        vaSpread1.Row = 0
        vaSpread1.FontName = "Tahoma"
        vaSpread1.FontSize = 12
        vaSpread1.Text = vaSpread1.Text
        vaSpread1.RowHeight(-1) = 23.5
        vaSpread1.RowHeight(0) = 23.5
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 15
        coladj = 7
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 14
          Next y
        Next x
        vaSpread1.RowHeight(0) = 24
        vaSpread1.RowHeight(-1) = 22
      Else
        COne = 6
        coladj = 2.3
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 11
          Next y
        Next x
        vaSpread1.RowHeight(0) = 19.5
        vaSpread1.RowHeight(-1) = 19.5
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 8
        coladj = 6
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 12
          Next y
        Next x
        vaSpread1.RowHeight(0) = 19.5
'        vaSpread1.FontBold = True
        vaSpread1.RowHeight(-1) = 19.5
      Else
        COne = 0.5
        coladj = 1.6
      End If
      Case 800
        COne = -0.6
        coladj = 1.55
        For x = 0 To 7
          For y = 0 To 2
            vaSpread1.FontName = "Tahoma"
            vaSpread1.Col = y
            vaSpread1.Row = x
            vaSpread1.FontSize = 10
          Next y
        Next x
        vaSpread1.RowHeight(0) = 14.75
        vaSpread1.RowHeight(-1) = 14.75
      Case Else
       
    End Select
    vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1) + COne
    vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + coladj
    vaSpread1.ColWidth(3) = 0
    vaSpread1.ColWidth(4) = 0

End Function

Private Sub ShowLaserLegal()
  Dim RptHandle As Integer
  Dim RptFile$
'  Dim TaxBill As TaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim x As Long, BillNo&
  Dim dlm$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  
  dlm$ = "~"
  
  RptFile$ = "TAXRPTS\TXLSRLEGAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
'  OpenTaxBillFile TBHandle, NumOfTBRecs
  '                              0                      1           2           3
  Print #RptHandle, CStr(TaxMasterRec.TaxYear); dlm; "1234"; dlm; "100"; dlm; "M45"; dlm;
  '                       4                5            6             7               8
  Print #RptHandle, "John Smith"; dlm; "5 acres"; dlm; "A"; dlm; "150,000.00"; dlm; "0.00"; dlm;
  '                    9              10              11           12            13            14
  Print #RptHandle, "0.00"; dlm; "150,000.00"; dlm; ".25"; dlm; "375.00"; dlm; "0.00"; dlm; "375.00"; dlm;
  '                         15                    16                      17
  Print #RptHandle, "Town of Anywhere"; dlm; "100 Main St"; dlm; "Anywhere, NC 55555"; dlm;
  '                     18                    19                      20                     21
  Print #RptHandle, "John Smith"; dlm; "500 Milltown Road"; dlm; "PO Box 1190"; dlm; "Anytown, NC 55555"
  
  Close
  
  arTaxLsrLegal.Show

End Sub
