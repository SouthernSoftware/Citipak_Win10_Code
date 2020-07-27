VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAEditItemWTabs 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset Edit Item"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAEditItemWTabs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   5775
      Left            =   240
      TabIndex        =   33
      Top             =   1950
      Width           =   11190
      _Version        =   196609
      _ExtentX        =   19748
      _ExtentY        =   10181
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   400
      TabsPerRow      =   3
      TabCount        =   3
      AlignTextH      =   1
      AlignTextV      =   1
      ThreeD          =   -1  'True
      GrayAreaColor   =   12632256
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      PageEarMarkColorDark=   12632256
      DataFormat      =   ""
      AutoSizeChildren=   2
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      DrawFocusRect   =   2
      DataField       =   ""
      TabCaption      =   "frmFAEditItemWTabs.frx":08CA
      PageEarMarkPictureNext=   "frmFAEditItemWTabs.frx":0C19
      PageEarMarkPicturePrev=   "frmFAEditItemWTabs.frx":0C35
      EarMarkPictureNext=   "frmFAEditItemWTabs.frx":0C51
      EarMarkPicturePrev=   "frmFAEditItemWTabs.frx":0C6D
      Begin VB.Frame Frame3 
         BackColor       =   &H008F8265&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Enabled         =   0   'False
         Height          =   4635
         Left            =   -25845
         TabIndex        =   52
         Top             =   -20400
         Width           =   10455
         Begin LpLib.fpCombo fpcmbDepYN 
            Height          =   405
            Left            =   5235
            TabIndex        =   20
            ToolTipText     =   "Enter a Y if this fixed asset will be depreciated or N if this fixed asset will not be depreciated (required)."
            Top             =   150
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
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
            ColDesigner     =   "frmFAEditItemWTabs.frx":0C89
         End
         Begin EditLib.fpText fptxtAssetLife 
            Height          =   396
            Left            =   8448
            TabIndex        =   25
            ToolTipText     =   "Enter the expected number of years this asset should be of value."
            Top             =   1536
            Width           =   588
            _Version        =   196608
            _ExtentX        =   1037
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
            CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
            MaxLength       =   150
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
         Begin EditLib.fpText fptxtLeft 
            Height          =   396
            Left            =   9264
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "This field displays the life remaining for this fixed asset."
            Top             =   1536
            Width           =   588
            _Version        =   196608
            _ExtentX        =   1037
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
            InvalidColor    =   12648447
            InvalidOption   =   0
            MarginLeft      =   3
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
            CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
            MaxLength       =   150
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
         Begin EditLib.fpCurrency fptxtOriginalCost 
            Height          =   396
            Left            =   2592
            TabIndex        =   22
            ToolTipText     =   "Enter the amount paid for this fixed asset (required)."
            Top             =   960
            Width           =   2508
            _Version        =   196608
            _ExtentX        =   4424
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
         Begin EditLib.fpText fptxtPONum 
            Height          =   396
            Left            =   7344
            TabIndex        =   23
            ToolTipText     =   "Enter the purchase order number for this fixed asset here (optional)."
            Top             =   960
            Width           =   2508
            _Version        =   196608
            _ExtentX        =   4424
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
         Begin EditLib.fpText fptxtChkNum 
            Height          =   396
            Left            =   2592
            TabIndex        =   24
            ToolTipText     =   "Enter the check number used to pay for this fixed asset here (optional)."
            Top             =   1536
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
         Begin EditLib.fpCurrency fptxtDep2Date 
            Height          =   396
            Left            =   2400
            TabIndex        =   29
            ToolTipText     =   "This field is automatically calculated when this fixed asset is depreciated. This value can be edited (but not recommended.)"
            Top             =   3024
            Width           =   2508
            _Version        =   196608
            _ExtentX        =   4424
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
         Begin EditLib.fpCurrency fptxtCurrVal 
            Height          =   396
            Left            =   7344
            TabIndex        =   30
            ToolTipText     =   "This value is automatically calculated. It cannot be edited."
            Top             =   2976
            Width           =   2508
            _Version        =   196608
            _ExtentX        =   4424
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
         Begin EditLib.fpDateTime fptxtCurrDepDate 
            Height          =   348
            Left            =   3168
            TabIndex        =   27
            ToolTipText     =   "This date, the most current depreciation date, is automatically calculated and cannot be edited."
            Top             =   2448
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
            ButtonStyle     =   2
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   12648447
            InvalidOption   =   0
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
            Text            =   "11/20/2002"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/dd/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
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
         Begin EditLib.fpDateTime fptxtEOLDate 
            Height          =   348
            Left            =   8112
            TabIndex        =   28
            ToolTipText     =   "This date, the End Of Life date, is automatically calculated and cannot be edited."
            Top             =   2448
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
            ButtonStyle     =   2
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
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
            Text            =   "11/20/2002"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/dd/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
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
         Begin EditLib.fpDateTime fptxtDisposalDate 
            Height          =   348
            Left            =   8064
            TabIndex        =   32
            ToolTipText     =   "This date indicates the date this item was disposed of and is figured in the disposal process. It is not editable here."
            Top             =   4032
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
            ButtonStyle     =   2
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   0   'False
            InvalidColor    =   12648447
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   0
            ControlType     =   1
            Text            =   "01/14/2003"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/dd/yyyy"
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
            ThreeDFrameColor=   -2147483633
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
         Begin EditLib.fpCurrency fptxtDispPrice 
            Height          =   396
            Left            =   2352
            TabIndex        =   31
            ToolTipText     =   "This value is calculated in the disposal process and is not editable here."
            Top             =   3984
            Width           =   2508
            _Version        =   196608
            _ExtentX        =   4424
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
         Begin EditLib.fpDateTime fptxtAcquiredDate 
            Height          =   348
            Left            =   8016
            TabIndex        =   21
            ToolTipText     =   "Enter the date on which this  fixed asset was purchased (required)."
            Top             =   144
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
            ButtonStyle     =   2
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "01/01/2003"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/dd/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
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
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Acquired On*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   6384
            TabIndex        =   73
            Top             =   192
            Width           =   1548
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   3
            X1              =   0
            X2              =   10464
            Y1              =   3696
            Y2              =   3696
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   3
            X1              =   0
            X2              =   10464
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0080FFFF&
            BorderWidth     =   3
            X1              =   0
            X2              =   10464
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   9072
            TabIndex        =   64
            Top             =   1584
            Width           =   156
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Disposal Price:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   528
            TabIndex        =   63
            Top             =   4080
            Width           =   1644
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Disposal Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   6288
            TabIndex        =   62
            Top             =   4080
            Width           =   1596
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "End Of Life Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   6000
            TabIndex        =   61
            Top             =   2496
            Width           =   1932
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Last Depreciation Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   384
            TabIndex        =   60
            Top             =   2496
            Width           =   2604
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Current Value:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5472
            TabIndex        =   59
            Top             =   3072
            Width           =   1692
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Deprec To Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   348
            Left            =   336
            TabIndex        =   58
            Top             =   3120
            Width           =   1884
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   384
            TabIndex        =   57
            Top             =   1632
            Width           =   2028
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "P.O. Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5232
            TabIndex        =   56
            Top             =   1056
            Width           =   2028
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Price*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   528
            TabIndex        =   55
            Top             =   1056
            Width           =   1884
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Do You Wish To Depreciate This Asset*?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   528
            TabIndex        =   54
            Top             =   240
            Width           =   4476
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Life Expectancy*/Life Left (Years):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   4512
            TabIndex        =   53
            Top             =   1584
            Width           =   3804
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H008F8265&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Enabled         =   0   'False
         Height          =   4635
         Left            =   -25845
         TabIndex        =   51
         Top             =   -20400
         Width           =   10455
         Begin EditLib.fpText fptxtVhclMake 
            Height          =   396
            Left            =   3852
            TabIndex        =   15
            ToolTipText     =   "Enter the make of this vehicle."
            Top             =   624
            Width           =   3468
            _Version        =   196608
            _ExtentX        =   6117
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
         Begin EditLib.fpText fptxtVhclModl 
            Height          =   396
            Left            =   3840
            TabIndex        =   16
            ToolTipText     =   "Enter the model of this vehicle."
            Top             =   1248
            Width           =   3468
            _Version        =   196608
            _ExtentX        =   6117
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
         Begin EditLib.fpText fptxtVIN 
            Height          =   396
            Left            =   3840
            TabIndex        =   17
            ToolTipText     =   "Enter the Vehicle Identification Number for this vehicle."
            Top             =   1872
            Width           =   3468
            _Version        =   196608
            _ExtentX        =   6117
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
         Begin EditLib.fpText fptxtLicNum 
            Height          =   396
            Left            =   3840
            TabIndex        =   18
            ToolTipText     =   "Enter the license number of this vehicle."
            Top             =   2496
            Width           =   3468
            _Version        =   196608
            _ExtentX        =   6117
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
         Begin EditLib.fpText fptxtVhclColr 
            Height          =   396
            Left            =   3840
            TabIndex        =   19
            ToolTipText     =   "Enter the color of this vehicle."
            Top             =   3120
            Width           =   3468
            _Version        =   196608
            _ExtentX        =   6117
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
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "These fields are designed to expand the available identifiers specific to vehicles. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   252
            Left            =   1680
            TabIndex        =   70
            Top             =   3840
            Width           =   7212
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Color:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   1824
            TabIndex        =   69
            Top             =   3216
            Width           =   1692
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "License Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   1392
            TabIndex        =   68
            Top             =   2592
            Width           =   2124
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " VIN:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   1824
            TabIndex        =   67
            Top             =   1968
            Width           =   1692
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " Model:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   1824
            TabIndex        =   66
            Top             =   1344
            Width           =   1692
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " Make:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   1836
            TabIndex        =   65
            Top             =   720
            Width           =   1692
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H008F8265&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4635
         Left            =   390
         TabIndex        =   36
         Top             =   765
         Width           =   10455
         Begin LpLib.fpCombo fpcmbStatus 
            Height          =   405
            Left            =   1965
            TabIndex        =   3
            ToolTipText     =   "Enter the active status of this fixed asset (required)."
            Top             =   1830
            Width           =   3135
            _Version        =   196608
            _ExtentX        =   5530
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
            ColDesigner     =   "frmFAEditItemWTabs.frx":0F80
         End
         Begin EditLib.fpText fptxtTagNumber 
            Height          =   396
            Left            =   1956
            TabIndex        =   0
            ToolTipText     =   "Enter the tag number here."
            Top             =   240
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
            CharValidationText=   "0 1 2 3 4 5 6 7 8 9 - "
            MaxLength       =   20
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
         Begin EditLib.fpText fptxtFundNum 
            Height          =   396
            Left            =   1956
            TabIndex        =   1
            ToolTipText     =   "Enter the General Ledger fund number here (required)."
            Top             =   768
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
            CharValidationText=   "0 , 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9"
            MaxLength       =   150
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
         Begin EditLib.fpText fptxtGLNum 
            Height          =   396
            Left            =   1968
            TabIndex        =   2
            ToolTipText     =   "Enter the desired general ledger number here (optional)."
            Top             =   1296
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - "
            MaxLength       =   14
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
         Begin EditLib.fpText fptxtDesc1 
            Height          =   396
            Left            =   1968
            TabIndex        =   4
            ToolTipText     =   "Enter a brief description of this fixed asset here (required)."
            Top             =   2352
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpText fptxtDesc2 
            Height          =   396
            Left            =   1968
            TabIndex        =   5
            ToolTipText     =   "Enter a brief description for this fixed asset here (optional)."
            Top             =   2880
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpText fptxtDeptNum 
            Height          =   396
            Left            =   1968
            TabIndex        =   6
            ToolTipText     =   "Enter a valid department number here (required)."
            Top             =   3408
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
         Begin EditLib.fpText fptxtGroupCode 
            Height          =   396
            Left            =   1968
            TabIndex        =   7
            ToolTipText     =   "Enter a valid asset code here (required)."
            Top             =   3936
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
            MaxLength       =   4
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
         Begin EditLib.fpText fptxtVendorNum 
            Height          =   396
            Left            =   7056
            TabIndex        =   8
            ToolTipText     =   "Enter the vendor of this fixed asset here (optional)."
            Top             =   240
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpText fptxtSerialNum 
            Height          =   396
            Left            =   7056
            TabIndex        =   9
            ToolTipText     =   "Enter the serial number for this fixed asset here (optional)."
            Top             =   768
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpText fptxtMfg 
            Height          =   396
            Left            =   7056
            TabIndex        =   10
            ToolTipText     =   "Enter the manufacturer of this fixed asset here (optional)."
            Top             =   1296
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpText fptxtContact 
            Height          =   396
            Left            =   7056
            TabIndex        =   11
            ToolTipText     =   "Enter the contact person for this fixed asset here (optional)."
            Top             =   1824
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpDateTime fpDateWrntyX 
            Height          =   348
            Left            =   7776
            TabIndex        =   13
            ToolTipText     =   "Enter the expiration date for the warranty for this fixed asset (optional)."
            Top             =   2880
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
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
            ButtonStyle     =   2
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   12648447
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "01/01/2003"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/dd/yyyy"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
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
         Begin EditLib.fpText fptxtLocation 
            Height          =   396
            Left            =   7056
            TabIndex        =   14
            ToolTipText     =   "Enter the location where this fixed asset can be found (optional)."
            Top             =   3936
            Width           =   3132
            _Version        =   196608
            _ExtentX        =   5524
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
         Begin EditLib.fpMask fptxtPhone 
            Height          =   396
            Left            =   7056
            TabIndex        =   12
            ToolTipText     =   "Enter the manufactuer's contact  phone number."
            Top             =   2352
            Width           =   2052
            _Version        =   196608
            _ExtentX        =   3619
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
            InvalidColor    =   -2147483643
            InvalidOption   =   0
            MarginLeft      =   3
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
         Begin fpBtnAtlLibCtl.fpBtn cmdFundList 
            Height          =   405
            Left            =   3930
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Click this button to bring up a list of all fund numbers."
            Top             =   765
            Width           =   1545
            _Version        =   131072
            _ExtentX        =   2725
            _ExtentY        =   714
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   13684944
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
            ButtonDesigner  =   "frmFAEditItemWTabs.frx":1277
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdTagList 
            Height          =   390
            Left            =   3936
            TabIndex        =   77
            TabStop         =   0   'False
            ToolTipText     =   "Click this button to bring up a list of all fixed assets."
            Top             =   240
            Width           =   1545
            _Version        =   131072
            _ExtentX        =   2725
            _ExtentY        =   688
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   13684944
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
            ButtonDesigner  =   "frmFAEditItemWTabs.frx":1457
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdDept 
            Height          =   390
            Left            =   3936
            TabIndex        =   78
            TabStop         =   0   'False
            ToolTipText     =   "Click this button to bring up a list of all fixed assets."
            Top             =   3408
            Width           =   1545
            _Version        =   131072
            _ExtentX        =   2725
            _ExtentY        =   688
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   13684944
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
            ButtonDesigner  =   "frmFAEditItemWTabs.frx":1637
         End
         Begin fpBtnAtlLibCtl.fpBtn cmdAssetList 
            Height          =   390
            Left            =   3936
            TabIndex        =   79
            TabStop         =   0   'False
            ToolTipText     =   "Click this button to bring up a list of all asset codes."
            Top             =   3936
            Width           =   1545
            _Version        =   131072
            _ExtentX        =   2725
            _ExtentY        =   688
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   13684944
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
            ButtonDesigner  =   "frmFAEditItemWTabs.frx":1817
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5856
            TabIndex        =   72
            Top             =   2448
            Width           =   1020
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Location:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5808
            TabIndex        =   50
            Top             =   4032
            Width           =   1068
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Warranty Expires On:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5184
            TabIndex        =   49
            Top             =   2928
            Width           =   2460
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5856
            TabIndex        =   48
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Manufacturer:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5328
            TabIndex        =   47
            Top             =   1392
            Width           =   1548
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Num:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5520
            TabIndex        =   46
            Top             =   864
            Width           =   1356
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   5712
            TabIndex        =   45
            Top             =   348
            Width           =   1212
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Group Code*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   240
            TabIndex        =   44
            Top             =   4032
            Width           =   1548
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dept. Num*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   288
            TabIndex        =   43
            Top             =   3504
            Width           =   1500
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   384
            TabIndex        =   42
            Top             =   2976
            Width           =   1404
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   336
            TabIndex        =   41
            Top             =   2448
            Width           =   1452
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   864
            TabIndex        =   40
            Top             =   1920
            Width           =   924
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "G/L Acct:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   528
            TabIndex        =   39
            Top             =   1392
            Width           =   1260
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "GL Fund*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   672
            TabIndex        =   38
            Top             =   864
            Width           =   1116
         End
         Begin VB.Label lblDesc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tag Number*:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   300
            Left            =   192
            TabIndex        =   37
            Top             =   348
            Width           =   1584
         End
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   4725
         Left            =   -25890
         Top             =   -20445
         Width           =   10545
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   4725
         Left            =   -25890
         Top             =   -20445
         Width           =   10545
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   4725
         Left            =   345
         Top             =   720
         Width           =   10545
      End
   End
   Begin EditLib.fpText fptxtHeader 
      Height          =   495
      Left            =   2460
      TabIndex        =   74
      ToolTipText     =   "Enter the vendor of this fixed asset here (optional)."
      Top             =   840
      Width           =   6735
      _Version        =   196608
      _ExtentX        =   11880
      _ExtentY        =   873
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
      ControlType     =   1
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   585
      Left            =   3312
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to return to the Item Lookup screen."
      Top             =   7875
      Width           =   2310
      _Version        =   131072
      _ExtentX        =   4075
      _ExtentY        =   1032
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAEditItemWTabs.frx":19F8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   585
      Left            =   6048
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to save all data entered above."
      Top             =   7875
      Width           =   2310
      _Version        =   131072
      _ExtentX        =   4075
      _ExtentY        =   1032
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAEditItemWTabs.frx":1BD4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTab 
      Height          =   585
      Left            =   8784
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Navigate tabs by pressing this button."
      Top             =   7875
      Width           =   2310
      _Version        =   131072
      _ExtentX        =   4075
      _ExtentY        =   1032
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFAEditItemWTabs.frx":1DB0
   End
   Begin VB.Label lblDspl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   4800
      TabIndex        =   71
      Top             =   1515
      Width           =   6540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2940
      TabIndex        =   35
      Top             =   330
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1230
      Index           =   1
      Left            =   1500
      Top             =   195
      Width           =   8655
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Required fields denoted with *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   195
      TabIndex        =   34
      Top             =   1560
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   1500
      Top             =   150
      Width           =   8655
   End
End
Attribute VB_Name = "frmFAEditItemWTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim BadGLNum As Boolean
  Dim BadAssetCodeNum As Boolean
  Dim BadTagNum As Boolean
  Dim FirstTagNum$
  Dim AssLife As Integer
  Dim AssLifeLeft As Integer
  Dim AcqDate As Integer
  Dim TempTagNumber$
  Dim TempItemTag$
  Dim TempISTATUS$
  Dim TempDEPYN$
  Dim TempAQURDATE As Integer
  Dim TempIDESC1$
  Dim TempIDESC2$
  Dim TempGLACCT$
  Dim TempIDEPT    As Integer
  Dim TempASSETCode$
  Dim TempILIFE    As Double
  Dim TempORGCOST  As Double
  Dim TempDEP2DATE As Double
  Dim TempCURRVAL  As Double
  Dim TempCDEPDATE As Integer
  Dim TempDispDate As Integer
  Dim TempVENDOR$
  Dim TempSERIALNO$
  Dim TempITEMMFG$
  Dim TempCONTACT$
  Dim TempPhone$
  Dim TempITEMLOC$
  Dim TempEOLDATE As Integer
  Dim TempVHCLMAKE$
  Dim TempVHCLMODL$
  Dim TempVHCLVIN$
  Dim TempVHCLTAG$
  Dim TempVHCLCOLR$
  Dim TempWARRXDAT As Integer
  Dim TempFundNum As Integer
  Dim TempDisposAmt As Double
  Dim TempLastDprRec As Long
  Dim TempLifeLeft As Integer
  Dim TempPONum$
  Dim TempCheckNum$
  Dim DsplFlag$
  Dim TempDsplMethod$
  Dim FirstTime As Boolean
  Dim EditFlag As Boolean
  Dim TempGRecNum As Integer
  
Private Sub cmdAssetList_Click()
  frmFAAssetCodeList.Show vbModal
End Sub

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal
End Sub

Public Sub cmdExit_Click()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  Dim TVHandle As Integer
  Dim TVRec As TempVHCLDataType
  Dim TVCnt As Integer
  Dim VChangeFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  
  ItemChangeFlag = False
  If VhclTempDsplFlag = True Then GoTo RecNumIsZero 'this means
  'that the fixed asset data on the screen represents a disposed
  'fixed asset...checking changes is not necessary (fields are
  'disabled for disposed of items)
  
  ChangeFlag = False
  
  If GRecNum = 0 Then  'user is exiting
    If MsgBox("Are you sure you want to proceed without saving any changes?", vbYesNo) = vbYes Then
  'without saving new record entries...also skips the change
  'check feature and if taglist is open then the number
  'double clicked will be brought up to this screen
      GoTo RecNumIsZero
    Else
      Close
      vaTabPro1.ActiveTab = 0
      fptxtTagNumber.SetFocus
      Exit Sub
    End If
  End If
  
  OpenFAItemFile FAHandle
  Get FAHandle, GRecNum, FAItemRec
  Close FAHandle
  
  If QPTrim$(fpDateWrntyX.Text) <> "NOT SAVED" Then
    If FAItemRec.WARRXDAT <> Date2Num(fpDateWrntyX) Then
      ChangeFlag = True
      vaTabPro1.ActiveTab = 0
      fpDateWrntyX.SetFocus
      GoTo ChangeFound
    End If
  End If
  
  If QPTrim$(FAItemRec.VENDOR) <> QPTrim$(fptxtVendorNum) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtVendorNum.SetFocus
    GoTo ChangeFound
  End If
  
  If Mid(fpcmbStatus.Text, 1, 1) <> QPTrim$(FAItemRec.ISTATUS) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbStatus.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.CONTACT) <> QPTrim$(fptxtContact.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtContact.SetFocus
    GoTo ChangeFound
  End If
    
  If QPTrim$(FAItemRec.Phone) <> QPTrim$(fptxtPhone) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtPhone.SetFocus
    GoTo ChangeFound
  End If
    
  If FAItemRec.IDEPT <> fptxtDeptNum Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtDeptNum.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(FAItemRec.IDESC1) <> QPTrim$(fptxtDesc1.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtDesc1.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.IDESC2) <> QPTrim$(fptxtDesc2.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtDesc2.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.GLACCT) <> QPTrim$(fptxtGLNum.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtGLNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ASSETCODE) <> QPTrim$(fptxtGroupCode.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtGroupCode.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.FundNum <> Val(QPTrim$(fptxtFundNum.Text)) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtFundNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ITEMLOC) <> QPTrim$(fptxtLocation.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtLocation.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ITEMMFG) <> QPTrim$(fptxtMfg.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtMfg.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.SERIALNO) <> QPTrim$(fptxtSerialNum.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSerialNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.ItemTag) <> QPTrim$(fptxtTagNumber.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtTagNumber.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.VHCLMAKE) <> QPTrim$(fptxtVhclMake.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtVhclMake.SetFocus
  End If
  
  If QPTrim$(FAItemRec.VHCLMODL) <> QPTrim$(fptxtVhclModl.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtVhclModl.SetFocus
  End If
  
  If QPTrim$(FAItemRec.VHCLVIN) <> QPTrim$(fptxtVIN.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtVIN.SetFocus
  End If
  
  If QPTrim$(FAItemRec.VHCLTAG) <> QPTrim$(fptxtLicNum.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtLicNum.SetFocus
  End If
  
  If QPTrim$(FAItemRec.VHCLCOLR) <> QPTrim$(fptxtVhclColr.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtVhclColr.SetFocus
  End If
  
  If QPTrim$(FAItemRec.DEPYN) <> QPTrim$(fpcmbDepYN.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fpcmbDepYN.SetFocus
    GoTo ChangeFound
  End If
  
  If FAItemRec.AQURDATE <> Date2Num(fptxtAcquiredDate) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtAcquiredDate.SetFocus
    GoTo ChangeFound
  End If
    
  If FAItemRec.ILIFE <> Val(fptxtAssetLife.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtAssetLife.SetFocus
    GoTo ChangeFound
  End If

  If OldRound(FAItemRec.DEP2DATE) <> OldRound(fptxtDep2Date.DoubleValue) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtDep2Date.SetFocus
    GoTo ChangeFound
  End If
    
  If FAItemRec.ORGCOST <> fptxtOriginalCost Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtOriginalCost.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.CheckNum) <> QPTrim$(fptxtChkNum.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtChkNum.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(FAItemRec.PONum) <> QPTrim$(fptxtPONum.Text) Then
    ChangeFlag = True
    vaTabPro1.ActiveTab = 2
    fptxtPONum.SetFocus
  End If
  
ChangeFound:
  'ItemChangeFlag..This flag is read by the tag list form to
  'know what to do with the decision made by the user below
  '-----------------------------------------------------------
  'The tag list can be double clicked to change the data on this screen. However, if
  'a change has been made before a new tag selection has been made and the user didn't
  'save it then he might lose data he thought was saved. So the ChangeFound routine looks
  'to see if the tag list is open ("taglistopen.dat") and if it is then we know it came
  'from the double click sub on that form where the .dat file is created. If the user wants
  'to abandon the changed data then the routine goes ahead and pops the screen with the
  'new tag data. If the user wants to save the changed data then the routine checks for
  'any save traps and if an error is found then the routine discards the new tag data and
  'returns the user to the screen to correct the error. If the user wants to save the changes
  'and there are no errors then the routine saves the data and pops the screen with the
  'new tag data. If the user wants to review any change then the routine discards the new
  'tag data and returns the user to the screen. The .dat file is always deleted so the tag
  'list can be reopened from scratch.
  If ChangeFlag = True Then
    ChangeFlag = False
    ItemChangeFlag = True
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
      Exit Sub 'don't exit
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      If Exist("taglistopen.dat") Then 'if the tag list is still open then
      'unload it and return to the screen with the focus on tag number (set above)
        Unload frmFATagList
        KillFile ("taglistopen.dat")
      End If
      Exit Sub 'go back to the screen
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      If Exist("taglistopen.dat") Then
        ItemChangeFlag = False 'tell the tag list that it's OK to continue
        'with changing the data on this screen with the new tag number entered
        KillFile ("taglistopen.dat")
        Exit Sub
      End If
      GRecNum = 0
      AddItemFlag = False
      frmFAItemLookUp.Show
      DoEvents
      KillFile ("edititemopen.dat")
      Unload frmFAEditItemWTabs
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
RecNumIsZero:
  If Exist("taglistopen.dat") Then 'without this the program would
  'exit out to the main menu when tag list was double clicked
    KillFile ("taglistopen.dat")
    Exit Sub
  End If
  
  If AddItemFlag = False Then
    frmFAItemLookUp.Show
  Else
    AddItemFlag = False
    frmFAItemMaintMenu.Show
  End If
  Close
  DoEvents
  KillFile ("edititemopen.dat")
  Unload frmFAEditItemWTabs
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "cmdExit_Click", Erl)
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
    Unload Me
  
End Sub

Private Sub cmdFundList_Click()
  frmFAFundList.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfRecs As Long
  Dim IdxFlag As Boolean
  Dim TVHandle As Integer
  Dim TVRec As TempVHCLDataType
  Dim TVCnt As Integer
  Dim TagChangeIndexFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  
  TagChangeIndexFlag = False
  IdxFlag = False
  'check for duplicate tag numbers only if the current tag entered
  'has been changed from the original number
  
  If QPTrim$(fptxtTagNumber.Text) <> FirstTagNum Then
    CheckForValidTAGNum
    If BadTagNum = True Then
      vaTabPro1.ActiveTab = 0
      fptxtTagNumber.SetFocus
      BadTagNum = False
      Exit Sub
    Else
      TagChangeIndexFlag = True 'we're changing a tag number so it will need indexing
    End If
  End If
  
  CheckForValidAssetCodeNum
  'when taglistopen.dat exists then we know the user has accessed it
  'to get another tag number
  If BadAssetCodeNum = True Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    BadAssetCodeNum = False
    Exit Sub
  End If
  
  If Check4ValidDept = False Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    Exit Sub
  End If
  
  If Check4ValidFund = False Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    Exit Sub
  End If
  
  OpenFAItemFile FAHandle
  NumOfRecs = LOF(FAHandle) \ Len(FAItemRec)
  
  If GRecNum = 0 Then 'new item here
    IdxFlag = True 'tells the program that tags need to be
    'reindexed to include this new item
    GRecNum = NumOfRecs + 1 'set record number where this data will
    'be saved
  Else
    Get FAHandle, GRecNum, FAItemRec 'otherwise pull up data for this asset
  End If
  
  If Len(QPTrim$(fpcmbDepYN.Text)) = 0 Then 'user forgot to fill in
  'the DepYN field
    If Exist("taglistopen.dat") Then 'OK...so if the tag list has been
    'accessed and is still open then the temp .dat file for that form should be deleted.
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no Y or N saved for depreciation. Please enter a Y or N for depreciation."
    Close FAHandle 'go back to the screen so the user can enter a value for DepYN
    If IdxFlag = True Then GRecNum = 0 'if this is a new record then reset the global to zero
    vaTabPro1.ActiveTab = 2
    fpcmbDepYN.SetFocus
    Exit Sub
  Else
    FAItemRec.DEPYN = QPTrim$(fpcmbDepYN.Text)
  End If
  
  If Len(QPTrim$(fpcmbStatus.Text)) = 0 Then 'nothing saved for this
  'required field so close temp files and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no date saved for this item's status. Please enter a status value."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fpcmbStatus.SetFocus
    Exit Sub
  ElseIf fpcmbStatus.Text = "Active" Then 'field OK so save value
    FAItemRec.ISTATUS = "A"
  Else
    FAItemRec.ISTATUS = "I"
  End If
  
  If Len(QPTrim(fptxtAcquiredDate)) = 0 Then 'nothing saved for this
  'required field so close temp files down and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no date saved for this item's acquired date. Please enter an acquired date."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtAcquiredDate.SetFocus
    Exit Sub
  Else 'else save this value
    FAItemRec.AQURDATE = Date2Num(fptxtAcquiredDate)
  End If
  
  If QPTrim$(fpDateWrntyX.Text) <> "NOT SAVED" Then
    'user entered a warranty date but it comes before when the
    'asset was purchased
    If Date2Num(fpDateWrntyX) < Date2Num(fptxtAcquiredDate) Then
      If MsgBox("The warranty date entered comes before the acquire date. Do you wish to continue anyway?", vbYesNo) = vbNo Then
        If Exist("taglistopen.dat") Then
          Unload frmFATagList
          KillFile ("taglistopen.dat")
        End If
        Close
        vaTabPro1.ActiveTab = 0
        fpDateWrntyX.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  If QPTrim$(fpDateWrntyX.Text) = "NOT SAVED" Then
    FAItemRec.WARRXDAT = 0
  Else
    FAItemRec.WARRXDAT = Date2Num(fpDateWrntyX)
  End If
  
  If fptxtAssetLife = 0 Then 'nothing saved for this required field
  'so close temp files and go back to screen
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's asset life. Please enter a value for asset life."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 2
    fptxtAssetLife.SetFocus
    Exit Sub
  Else
    FAItemRec.ILIFE = Val(fptxtAssetLife) 'save this valid data
  End If
  
  FAItemRec.LifeLeft = Val(fptxtLeft.Text) 'can be edited but
  'is also figured automatically
  
  FAItemRec.CONTACT = QPTrim$(fptxtContact) 'not required
  FAItemRec.Phone = fptxtPhone
  If fptxtCurrDepDate = "NOT SAVED" Then
    FAItemRec.CDEPDATE = -11001 'value represents an invalid date...
    'program validates any date over -11000 (and under the year 2100)
  Else
    FAItemRec.CDEPDATE = Date2Num(fptxtCurrDepDate) 'save valid date
  End If
  
  FAItemRec.CURRVAL = fptxtCurrVal 'locked and automatically figured
  FAItemRec.DEP2DATE = fptxtDep2Date 'locked and automatically figured

  If Len(QPTrim(fptxtDeptNum)) = 0 Then 'required field with no valid value
    If Exist("taglistopen.dat") Then 'close down temp files and return to screen
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's department number. Please enter a value for department number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtDeptNum.SetFocus
    Exit Sub
  Else
    FAItemRec.IDEPT = Val(fptxtDeptNum) 'valid value so save it
  End If

  'no description entered for this item so close down temp files
  'and return to screen for correction
  If QPTrim$(fptxtDesc1) = "" And QPTrim$(fptxtDesc2) = "" Then
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "No description has been entered for this item. Please enter a description."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtDesc1.SetFocus
    Exit Sub
  Else 'otherwise save descriptions as entered
    FAItemRec.IDESC1 = QPTrim$(fptxtDesc1)
    FAItemRec.IDESC2 = QPTrim$(fptxtDesc2)
  End If
  
  'disposal data handling
  If fptxtDisposalDate = "NOT SAVED" Then
    FAItemRec.DsplFlag = 0
    FAItemRec.DispDate = 0
    GoTo NotDisposed
  ElseIf FAItemRec.DsplFlag = 1 Then
    GoTo NotDisposed
  Else
    FAItemRec.DispDate = Date2Num(fptxtDisposalDate) 'only if we allow disposal data to be changed here and not in the disposal routine
    FAItemRec.DsplFlag = 2
  End If
NotDisposed:
  FAItemRec.EOLDATE = Date2Num(fptxtEOLDate) 'a date will always be in here
  
  If Len(QPTrim(fptxtFundNum)) = 0 Then 'fund number is a required field
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's Fund number. Please enter a value for Fund number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtFundNum.SetFocus
    Exit Sub
  Else
    FAItemRec.FundNum = Val(QPTrim(fptxtFundNum)) 'save this valid fund number...
    'we know it's a valid fund number because it was checked earlier
  End If
  
  FAItemRec.GLACCT = QPTrim$(fptxtGLNum) 'not required
  
  If Len(QPTrim(fptxtGroupCode)) = 0 Then 'asset group code entry isn't valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's group code number. Please enter a value for group code number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtGroupCode.SetFocus
    Exit Sub
  Else
    FAItemRec.ASSETCODE = QPTrim$(fptxtGroupCode) 'value is good, it's been
    'checked in Check4ValidAssetCode...save it
  End If
  
  FAItemRec.ITEMLOC = QPTrim$(fptxtLocation) 'not required
  
  FAItemRec.ITEMMFG = QPTrim$(fptxtMfg) 'not required
  
  If fptxtOriginalCost = 0 Then 'purchase price value is not valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's purchase price. Please enter a value for purchase price."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 2
    fptxtOriginalCost.SetFocus
    Exit Sub
  Else
    FAItemRec.ORGCOST = fptxtOriginalCost 'purchase price is valid so save it
  End If
  
  FAItemRec.SERIALNO = QPTrim$(fptxtSerialNum) 'not required
  
  If Len(QPTrim(fptxtTagNumber)) = 0 Then 'essential tag number not valid
    If Exist("taglistopen.dat") Then
      Unload frmFATagList
      KillFile ("taglistopen.dat")
    End If
    MsgBox "There is no value saved for this item's tag number. Please enter a value for tag number."
    If IdxFlag = True Then GRecNum = 0
    Close FAHandle
    vaTabPro1.ActiveTab = 0
    fptxtTagNumber.SetFocus
    Exit Sub
  Else
    FAItemRec.ItemTag = QPTrim$(fptxtTagNumber) 'checked in Check4ValidTagNum and
    'is valid so save
  End If
  
  FAItemRec.VENDOR = QPTrim$(fptxtVendorNum) 'not required
  
  FAItemRec.CheckNum = QPTrim$(fptxtChkNum.Text) 'not required
  FAItemRec.PONum = QPTrim$(fptxtPONum.Text) 'not required
  
  FAItemRec.VHCLMAKE = QPTrim$(fptxtVhclMake.Text)
  FAItemRec.VHCLMODL = QPTrim$(fptxtVhclModl.Text)
  FAItemRec.VHCLVIN = QPTrim$(fptxtVIN.Text)
  FAItemRec.VHCLTAG = QPTrim$(fptxtLicNum.Text)
  FAItemRec.VHCLCOLR = QPTrim$(fptxtVhclColr.Text)
  
  If IdxFlag = True Then 'save values as empty just to hold space
    FAItemRec.Fill1 = ""
    FAItemRec.LastDprRec = 0
    FAItemRec.DsplMethod = ""
  End If
  
  Put FAHandle, GRecNum, FAItemRec 'save it to disk
  Close FAHandle
  
  If IdxFlag = True Or TagChangeIndexFlag = True Then 'this is a new asset so work it into
  'the tag index
    Call CreateTagIdx
    MainLog ("Item number " + QPTrim$(FAItemRec.ItemTag) + " data saved in frmFAEditItemWTabs.")
    IdxFlag = False
    TagChangeIndexFlag = False
  Else
    Call LogSaves 'records any changes made in this save with existing items
  End If
  
  MsgBox "Item data for " + QPTrim$(FAItemRec.ItemTag) + " has been saved."
  'If this save request was initiated by a double click from tag list then return
  'control back to that form

NoAsset:
  If Exist("taglistopen.dat") Then 'If a user has made a change and then
  'double clicked the tag list but did not save his change...he was alerted and
  'decide to save the change...this if statement sends the new tag data
  '(just double clicked) to the screen instead of exiting to the main menu
    KillFile ("taglistopen.dat")
    ItemChangeFlag = False
    vaTabPro1.ActiveTab = 0
    fptxtTagNumber.SetFocus
    Exit Sub
  End If
  
  If AddItemFlag = True Then 'entering a list of several items is tedious
  'if after each save the program returns to the menu so this feature allows
  'the user to speed up the entry process
    If MsgBox("Do you wish to add another new item?", vbYesNo) = vbYes Then
      GRecNum = 0
      vaTabPro1.ActiveTab = 0
      fptxtTagNumber.SetFocus
      Call LoadMe
      Exit Sub
    Else
      frmFAItemMaintMenu.Show
      DoEvents
      KillFile ("edititemopen.dat")
      Unload frmFAEditItemWTabs
      Exit Sub
    End If
  Else 'just editing an existing item...sends user back to menu upon
  'completion
    GRecNum = 0
    ThisTag = QPTrim$(fptxtTagNumber.Text)
    frmFAItemLookUp.Show
    Call frmFAItemLookUp.cmdSearch_Click
    DoEvents
    KillFile ("edititemopen.dat")
    Unload frmFAEditItemWTabs
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "cmdSave_Click", Erl)
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
    Unload Me
  
End Sub

Private Sub cmdTab_Click()
  On Error Resume Next
  If vaTabPro1.ActiveTab = 0 Then
    vaTabPro1.ActiveTab = 1
    If DsplFlag$ = True Then GoTo DisposedOf
    fptxtVhclMake.SetFocus
  ElseIf vaTabPro1.ActiveTab = 1 Then
    vaTabPro1.ActiveTab = 2
    If DsplFlag$ = True Then GoTo DisposedOf
    fpcmbDepYN.SetFocus
  ElseIf vaTabPro1.ActiveTab = 2 Then
    vaTabPro1.ActiveTab = 0
    If DsplFlag$ = True Then GoTo DisposedOf
    fptxtTagNumber.SetFocus
  End If
DisposedOf:
End Sub

Private Sub cmdTagList_Click()
  frmFATagList.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstTime = True
  EditFlag = False
  TempTagNumber = ""
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
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%N"
      Call cmdTab_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%F"
      Call cmdFundList_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      Call cmdDept_Click
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%T"
      Call cmdTagList_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%L"
      Call cmdAssetList_Click
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
      KillFile ("edititemopen.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditItemWTabs.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Public Sub LoadMe()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim Today As String * 10
  Dim One As Integer
  Dim FileHandle As Integer
  
  On Error GoTo ERRORSTUFF
  If ScreenW <> 1024 Then 'the tab fonts do not behave nicely
  'if the screen resolution is not 1024 so you have to fool it
  'into keeping the Tab 0 font from inflating when it shouldn't
    vaTabPro1.Tab = 1
    vaTabPro1.TabHeight = 400
    vaTabPro1.FontSize = 12
    vaTabPro1.Tab = 2
    vaTabPro1.TabHeight = 400
    vaTabPro1.FontSize = 12
    vaTabPro1.Tab = 0
    vaTabPro1.TabHeight = 400
    vaTabPro1.FontSize = 10
  End If
  
  lblDspl.Visible = False
  DsplFlag$ = False
  One = 1
  FileHandle = FreeFile
  Open "edititemopen.dat" For Output As FileHandle Len = 2
  
  Print #FileHandle, One
  Close FileHandle
  
  'Date$ = FormatDateTime(Date, vbShortDate)
  
  Today = FormatDateTime(Date, vbShortDate)
  
  fpcmbDepYN.Clear
  fpcmbDepYN.AddItem "Y"
  fpcmbDepYN.AddItem "N"
  
  fpcmbStatus.Clear
  fpcmbStatus.AddItem "Active"
  fpcmbStatus.AddItem "Inactive"
  
  If GRecNum = 0 Then 'load procedure for adding a new item
'    With AddItemFlag = True setting ...this global alerts the program
    'to use the "add another item" feature allowing the user to
    'speed up the entry process of a list of items to ass...keeps the
    'program from exiting out to the menu after each save
    fpcmbDepYN.Text = "N" 'defaults to this
    fpcmbStatus.Text = "Active" 'defaults to this
    fptxtAcquiredDate = Today 'defaults to this
    AcqDate = Date2Num(fptxtAcquiredDate)
    fptxtAssetLife = "0" 'defaults to this
    AssLife = 0 'defaults to this
    fptxtContact = "" 'defaults to this
    fptxtPhone = "000-000-0000"
    fptxtCurrDepDate = "NOT SAVED" 'defaults to this
    fptxtCurrVal = "0.00" 'defaults to this
    fptxtDep2Date = "0.00" 'defaults to this
    fptxtDeptNum = ""
    fptxtDesc1 = ""
    fptxtDesc2 = ""
    fptxtDisposalDate = "12/31/1979" 'means nothing is saved
    fptxtEOLDate = Today 'defaults to this
    fptxtGLNum = ""
    fptxtGroupCode = ""
    fptxtLocation = ""
    fptxtMfg = ""
    fptxtOriginalCost = "0.00" 'defaults to this
    fptxtSerialNum = "" 'etc
    fptxtFundNum = ""
    fptxtTagNumber = "" 'etc
    FirstTagNum = ""
    fptxtVendorNum = ""
    fptxtDispPrice = 0
    fptxtLeft.Text = 0
    fpDateWrntyX = "NOT SAVED" 'defaults to this
    fptxtPONum.Text = ""
    fptxtChkNum.Text = ""
    fptxtLicNum.Text = ""
    fptxtVhclModl.Text = ""
    fptxtVhclMake.Text = ""
    fptxtVIN.Text = ""
    fptxtVhclColr.Text = ""
  Else 'load procedure for an existing item
    OpenFAItemFile FAHandle
    Get FAHandle, GRecNum, FAItemRec
    If FAItemRec.DsplFlag > 0 Then 'this item is either disposed of or in the process
    'of being disposed of
      lblDspl.Visible = True 'show label telling of the disposal date
      DsplFlag$ = True
      If FAItemRec.DsplFlag = 2 Then 'if it's disposed of
        lblDspl.Caption = "This item was disposed of on " + MakeRegDate(FAItemRec.DispDate) + "."
      ElseIf FAItemRec.DsplFlag = 1 Then 'or it's in the disposal process
        lblDspl.Caption = "This item is scheduled for disposal on " + MakeRegDate(FAItemRec.DispDate) + "."
      End If
      'since this item is disposed of then disable the following
      cmdAssetList.Enabled = False
      cmdDept.Enabled = False
      cmdSave.Enabled = False
      cmdFundList.Enabled = False
      fpcmbDepYN.Enabled = False
      fpcmbStatus.Enabled = False
      fptxtAcquiredDate.Enabled = False
      fptxtAssetLife.Enabled = False
      fptxtContact.Enabled = False
      fptxtPhone.Enabled = False
      fptxtDep2Date.Enabled = False
      fptxtDeptNum.Enabled = False
      fptxtDesc1.Enabled = False
      fptxtDesc2.Enabled = False
      fpDateWrntyX.Enabled = False
      fptxtPONum.Enabled = False
      fptxtChkNum.Enabled = False
      fptxtGLNum.Enabled = False
      fptxtFundNum.Enabled = False
      fptxtGroupCode.Enabled = False
      fptxtOriginalCost.Enabled = False
      fptxtSerialNum.Enabled = False
      fptxtTagNumber.Enabled = False
      fptxtLocation.Enabled = False
      fptxtMfg.Enabled = False
      fptxtVendorNum.Enabled = False
      fptxtLeft.Enabled = False
      fptxtLicNum.Enabled = False
      fptxtVhclModl.Enabled = False
      fptxtVhclMake.Enabled = False
      fptxtVIN.Enabled = False
      fptxtVhclColr.Enabled = False
    Else 'this item is not disabled so activate the following
      cmdAssetList.Enabled = True
      cmdDept.Enabled = True
      cmdSave.Enabled = True
      cmdFundList.Enabled = True
      lblDspl.Visible = False
      fpcmbDepYN.Enabled = True
      fpcmbStatus.Enabled = True
      fptxtAcquiredDate.Enabled = True
      fptxtAssetLife.Enabled = True
      fptxtContact.Enabled = True
      fptxtPhone.Enabled = True
      fptxtDep2Date.Enabled = True
      fptxtDeptNum.Enabled = True
      fptxtDesc1.Enabled = True
      fptxtDesc2.Enabled = True
      fpDateWrntyX.Enabled = True
      fptxtPONum.Enabled = True
      fptxtChkNum.Enabled = True
      fptxtGLNum.Enabled = True
      fptxtFundNum.Enabled = True
      fptxtGroupCode.Enabled = True
      fptxtOriginalCost.Enabled = True
      fptxtSerialNum.Enabled = True
      fptxtTagNumber.Enabled = True
      fptxtLocation.Enabled = True
      fptxtMfg.Enabled = True
      fptxtVendorNum.Enabled = True
      fptxtLeft.Enabled = True
      fptxtLicNum.Enabled = True
      fptxtVhclModl.Enabled = True
      fptxtVhclMake.Enabled = True
      fptxtVIN.Enabled = True
      fptxtVhclColr.Enabled = True
    End If
    'now populate fields regardless if they are enabled or disabled
    '...globals are used when save routine occurs to record any changes
    'to main log
    FirstTagNum = QPTrim$(FAItemRec.ItemTag)
    TempItemTag$ = QPTrim$(FAItemRec.ItemTag) 'local global
    fpcmbDepYN.Text = FAItemRec.DEPYN
    TempDEPYN$ = FAItemRec.DEPYN 'local global
    If QPTrim$(FAItemRec.ISTATUS) = "A" Then
      fpcmbStatus.Text = "Active"
    Else
      fpcmbStatus.Text = "Inactive"
    End If
    TempISTATUS$ = QPTrim$(FAItemRec.ISTATUS) 'local global
    fptxtAcquiredDate = MakeRegDate(FAItemRec.AQURDATE)
    AcqDate = FAItemRec.AQURDATE
    TempAQURDATE = FAItemRec.AQURDATE 'local global
    If FAItemRec.ILIFE < 0 Then
      fptxtAssetLife = "0"
      AssLife = 0
    Else
      fptxtAssetLife = FAItemRec.ILIFE
      AssLife = FAItemRec.ILIFE
    End If
    TempILIFE = FAItemRec.ILIFE 'local global
    fptxtContact = QPTrim$(FAItemRec.CONTACT)
    TempCONTACT$ = QPTrim$(FAItemRec.CONTACT) 'local global
    fptxtPhone = FAItemRec.Phone
    TempPhone = FAItemRec.Phone
    If FAItemRec.CDEPDATE < -11000 Then 'roughly 1950
      FAItemRec.CDEPDATE = 0
      fptxtCurrDepDate = "NOT SAVED"
    Else
      fptxtCurrDepDate = MakeRegDate(FAItemRec.CDEPDATE)
    End If
    TempCDEPDATE = FAItemRec.CDEPDATE 'local global
    fptxtCurrVal = FAItemRec.CURRVAL '
    TempCURRVAL = FAItemRec.CURRVAL 'local global
    fptxtDep2Date = FAItemRec.DEP2DATE
    TempDEP2DATE = FAItemRec.DEP2DATE 'local global
    fptxtDeptNum = FAItemRec.IDEPT
    TempIDEPT = FAItemRec.IDEPT 'local global
    fptxtDesc1 = QPTrim$(FAItemRec.IDESC1)
    fptxtHeader.Text = QPTrim$(FAItemRec.ItemTag) + "  " + QPTrim$(FAItemRec.IDESC1)
    TempIDESC1$ = QPTrim$(FAItemRec.IDESC1) 'local global
    fptxtDesc2 = QPTrim$(FAItemRec.IDESC2)
    TempIDESC2$ = QPTrim$(FAItemRec.IDESC2) 'local global
    fptxtDisposalDate = MakeRegDate(FAItemRec.DispDate)
    TempDispDate = FAItemRec.DispDate 'local global
    If CheckValDate(fptxtDisposalDate) = False Then
      fptxtDisposalDate.Text = "NOT SAVED"
    ElseIf FAItemRec.DispDate = 0 Then
      fptxtDisposalDate.Text = "NOT SAVED"
    End If
    
    fpDateWrntyX = MakeRegDate(FAItemRec.WARRXDAT)
    
    If CheckValDate(fpDateWrntyX.Text) = False Then
      fpDateWrntyX.Text = "NOT SAVED"
    ElseIf FAItemRec.WARRXDAT = 0 Then
      fpDateWrntyX.Text = "NOT SAVED"
    End If
    TempWARRXDAT = FAItemRec.WARRXDAT 'local global
    fptxtPONum.Text = QPTrim$(FAItemRec.PONum)
    TempPONum$ = QPTrim$(FAItemRec.PONum) 'local global
    fptxtChkNum.Text = QPTrim$(FAItemRec.CheckNum)
    TempCheckNum$ = QPTrim$(FAItemRec.CheckNum) 'local global
    fptxtEOLDate = MakeRegDate(FAItemRec.EOLDATE)
    TempEOLDATE = FAItemRec.EOLDATE 'local global
    fptxtGLNum = QPTrim$(FAItemRec.GLACCT)
    TempGLACCT$ = QPTrim$(FAItemRec.GLACCT) 'local global
    fptxtFundNum = FAItemRec.FundNum
    TempFundNum = FAItemRec.FundNum 'local global
    fptxtGroupCode = QPTrim$(FAItemRec.ASSETCODE)
    TempASSETCode$ = QPTrim$(FAItemRec.ASSETCODE) 'local global
    fptxtLocation = QPTrim$(FAItemRec.ITEMLOC)
    TempITEMLOC$ = QPTrim$(FAItemRec.ITEMLOC) 'local global
    fptxtMfg = QPTrim$(FAItemRec.ITEMMFG)
    TempITEMMFG$ = QPTrim$(FAItemRec.ITEMMFG) 'local global
    fptxtOriginalCost = FAItemRec.ORGCOST
    TempORGCOST = FAItemRec.ORGCOST 'local global
    fptxtSerialNum = QPTrim$(FAItemRec.SERIALNO)
    TempSERIALNO$ = QPTrim$(FAItemRec.SERIALNO) 'local global
    fptxtTagNumber = QPTrim$(FAItemRec.ItemTag)
    fptxtVendorNum = QPTrim$(FAItemRec.VENDOR)
    TempVENDOR$ = QPTrim$(FAItemRec.VENDOR) 'local global
    fptxtDispPrice = FAItemRec.DisposAmt
    TempDisposAmt = FAItemRec.DisposAmt 'local global
    fptxtLeft.Text = FAItemRec.LifeLeft
    TempLifeLeft = FAItemRec.LifeLeft 'local global
    AssLifeLeft = FAItemRec.LifeLeft
    fptxtVhclMake.Text = QPTrim$(FAItemRec.VHCLMAKE)
    fptxtVhclModl.Text = QPTrim$(FAItemRec.VHCLMODL)
    fptxtVIN.Text = QPTrim$(FAItemRec.VHCLVIN)
    fptxtLicNum.Text = QPTrim$(FAItemRec.VHCLTAG)
    fptxtVhclColr.Text = QPTrim$(FAItemRec.VHCLCOLR)
    TempVHCLMAKE$ = QPTrim$(FAItemRec.VHCLMAKE)
    TempVHCLMODL$ = QPTrim$(FAItemRec.VHCLMODL)
    TempVHCLVIN$ = QPTrim$(FAItemRec.VHCLVIN)
    TempVHCLTAG$ = QPTrim$(FAItemRec.VHCLTAG)
    TempVHCLCOLR$ = QPTrim$(FAItemRec.VHCLCOLR)
    Close FAHandle
  End If
  If PWcnt = 0 Then 'this is a special case if sosoft needs
  'to access any field because of some kind of unexpected
  'problem...entering fixed assets with the sosoft code allows this
    fpcmbDepYN.Enabled = True
    fpcmbStatus.Enabled = True
    fptxtAcquiredDate.Enabled = True
    fptxtAssetLife.Enabled = True
    fptxtContact.Enabled = True
    fptxtPhone.Enabled = True
    fptxtDep2Date.Enabled = True
    fptxtDeptNum.Enabled = True
    fptxtDesc1.Enabled = True
    fptxtDesc2.Enabled = True
    fpDateWrntyX.Enabled = True
    fptxtPONum.Enabled = True
    fptxtChkNum.Enabled = True
    fptxtGLNum.Enabled = True
    fptxtFundNum.Enabled = True
    fptxtGroupCode.Enabled = True
    fptxtOriginalCost.Enabled = True
    fptxtSerialNum.Enabled = True
    fptxtTagNumber.Enabled = True
    fptxtLocation.Enabled = True
    fptxtMfg.Enabled = True
    fptxtLeft.Enabled = True
    fptxtLeft.ControlType = ControlTypeNormal
    fptxtCurrDepDate.ControlType = ControlTypeNormal
    fptxtCurrVal.ControlType = ControlTypeNormal
    fptxtDep2Date.ControlType = ControlTypeNormal
    fptxtDisposalDate.ControlType = ControlTypeNormal
    fptxtDispPrice.ControlType = ControlTypeNormal
    fptxtEOLDate.ControlType = ControlTypeNormal
    fptxtVhclMake.Enabled = True
    fptxtVhclModl.Enabled = True
    fptxtVIN.Enabled = True
    fptxtLicNum.Enabled = True
    fptxtVhclColr.Enabled = True
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "LoadMe", Erl)
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
    Unload Me
  
End Sub
Private Sub fpcmbDepYN_KeyDown(KeyCode As Integer, Shift As Integer)
  'This routine is designed to allow the user to scroll through the
  'form without inadvertently changing data in this combo box
  If KeyCode = vbKeySpace Then
    fpcmbDepYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDepYN.ListIndex = -1
  End If
  If fpcmbDepYN.ListDown <> True Then
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

Private Sub fpcmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  'This routine is designed to allow the user to scroll through the
  'form without inadvertently changing data in this combo box
  If KeyCode = vbKeySpace Then
    fpcmbStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStatus.ListIndex = -1
  End If
  If fpcmbStatus.ListDown <> True Then
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

Private Sub fptxtAcquiredDate_LostFocus()
  Dim AcqYear As Integer
  Dim NewEOL As Integer
  'this routine automatically updates the EOL field if
  'the user changes the acquire date
  On Error Resume Next
  
  If Date2Num(fptxtAcquiredDate) <> AcqDate Then 'program sees that this field has changed
  'from the currently saved global acquire date
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'find the new acquire year
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'now determine the new EOL Year
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL) 'now assign entire date to EOL
    AcqDate = Date2Num(fptxtAcquiredDate) 'reassign global
  End If

End Sub

Private Sub fptxtAssetLife_Change()
  'goes ahead and figures the life left based on the
  'assigned life for this new asset
  If GRecNum = 0 Then
    fptxtLeft.Text = fptxtAssetLife.Text
  End If
  
End Sub

Private Sub fptxtAssetLife_LostFocus()
  Dim AcqYear As Integer
  Dim NewEOL As Integer
  Dim LifeDif As Integer
  
  On Error Resume Next
  'asset life affects when the assets EOL will be so this routine
  'determines EOL and asset life left
  If CInt(fptxtAssetLife.Text) <= 0 Then 'this may change but for now all
  'assets must have a life of at least 1 year
    MsgBox "Each fixed asset must have a life of at least 1 year."
    fptxtAssetLife = 1
  End If
  
  If GRecNum = 0 Then 'this is a new asset being added
    fptxtLeft.Text = fptxtAssetLife.Text 'life and life left are the
    'same for a new asset
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'assign global
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'now figure EOL
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL)
    If fptxtAssetLife <> AssLife Then 'update global if necessary
      AssLife = fptxtAssetLife
    End If
    Exit Sub
  End If
  
  If fptxtAssetLife = "" Then 'reassign global if field is blank
    fptxtAssetLife = AssLife
  End If
  
  If fptxtAssetLife <> AssLife Then 'change detected
    AcqYear = CInt(Mid(fptxtAcquiredDate, 7, 4)) 'start changing the EOL Date
    NewEOL = AcqYear + CInt(fptxtAssetLife) 'get new EOL Year
    fptxtEOLDate = Mid(fptxtAcquiredDate, 1, 6) + CStr(NewEOL) 'assign entire EOL Date
    If AssLife > CInt(fptxtAssetLife.Text) Then 'old asset life was more than the new one
      LifeDif = AssLife - CInt(fptxtAssetLife.Text) 'get difference between old and new
      fptxtLeft.Text = CInt(fptxtLeft.Text) - LifeDif 'new life left value is reduced to this value
      If Val(fptxtLeft.Text) < 0 Then fptxtLeft.Text = 0 'if the new life left ends up being
      'less than 0 then make it 0
    ElseIf AssLife < CInt(fptxtAssetLife.Text) Then 'old asset life is less than the new asset life
      LifeDif = CInt(fptxtAssetLife.Text) - AssLife 'get difference between new and old
      fptxtLeft.Text = CInt(fptxtLeft.Text) + LifeDif 'increase life left by the difference
      If Val(fptxtLeft.Text) < 0 Then fptxtLeft.Text = 0 'this should never happen
    End If
    AssLife = CInt(fptxtAssetLife) 'reassign global
    AssLifeLeft = CInt(fptxtLeft.Text)
  End If
  
End Sub

Private Sub fptxtDep2Date_LostFocus()
  Dim ORCost As Double
  Dim DepToDate As Double
  
  'this procedure updates the fixed asset's current value
  'if the user changes the depreciation amount to date manually
  ORCost = fptxtOriginalCost
  DepToDate = fptxtDep2Date
  fptxtCurrVal = ORCost - DepToDate

End Sub

Private Sub fptxtDesc1_LostFocus()
  fptxtHeader.Text = QPTrim$(fptxtTagNumber.Text) + "  " + QPTrim$(fptxtDesc1.Text)
End Sub

Private Sub fptxtDisposalDate_Change()

  If fptxtDisposalDate.Text = "12/31/1979" Or QPTrim$(fptxtDisposalDate) = "" Then
    fptxtDisposalDate.Text = "NOT SAVED"
    Exit Sub
  End If
  
  If fptxtDisposalDate.Text <> "NOT SAVED" And fptxtDispPrice > 0 Then
    fptxtCurrVal = 0 'if a disposal price and date have been saved then this
    'asset is no longer owned and the current value has to be zero
  End If
  
End Sub

Private Sub fptxtDisposalDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbDepYN.SetFocus
  End If
End Sub

Private Sub fptxtDispPrice_Change()
  If fptxtDisposalDate.Text <> "NOT SAVED" And fptxtDispPrice > 0 Then
    fptxtCurrVal = 0 'if the disposal price is more than 0 and the date is valid then the item is assumed
    'to be disposed of and has no more value
  End If

End Sub

Private Sub CheckForValidAssetCodeNum()
   Dim CodeRec As FAAssetCodeRecType
   Dim ACHandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim ThisText$
   
   'this routine is designed to make sure that the asset code entered
   'by the user is actually one of the codes saved
   On Error GoTo ERRORSTUFF
   
   ThisText$ = QPTrim$(fptxtGroupCode) 'user entered no value
   If ThisText$ = "" Then GoTo ZeroText 'exit sub...this is a required
   'field so if the user tries to save an empty filed the program
   'will force him to enter a valid number
   
   BadAssetCodeNum = False 'so far this asset code is OK
   
   OpenFACodeNameFile ACHandle
   TotalAccts = LOF(ACHandle) \ Len(CodeRec)
   
   If TotalAccts = 0 Then
     Close
     frmFAEditItemMess.cmdExit.Text = "ESC &Continue Saving"
     frmFAEditItemMess.cmdCont.Text = "F10 &Jump to Asset Code Edit"
     frmFAEditItemMess.Label1.Caption = "Fixed assets need to be categorized according to the asset code each fixed asset is assigned. Since no asset codes have been set up this fixed asset tracking feature is not possible. It is recommended that all asset codes are set up before you continue saving fixed assets. If you wish to jump to the asset code edit screen then press F10. Otherwise press ESC to return to the screen."
     frmFAEditItemMess.Show vbModal
     If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
       BadAssetCodeNum = True
       Unload frmFAEditItemMess
       frmFAEditAssetCode.Show
       DoEvents
       Unload frmFAItemLookUp
       Unload Me
       Exit Sub
     Else
       MainLog ("User warned that no asset code numbers have been saved. The user elected to save item data anyway for tag number " + fptxtTagNumber.Text + " in frmFAEditItemWTabs.")
       Unload frmFAEditItemMess
       Me.vaTabPro1.ActiveTab = 0
       fptxtGroupCode.SetFocus
       Exit Sub
     End If
   End If
   
   'go thru each number one at a time and compare against all asset code nums
   For x = 1 To TotalAccts
     Get ACHandle, x, CodeRec
       If ThisText = QPTrim$(CodeRec.ASSETCODE) Then 'found a match...this number is OK
         Exit For 'no reason to continue matching
       End If
  Next x
  
  If x = TotalAccts + 1 Then 'been thru all depts and found nothing to match
  'what the user entered
    Close
    BadAssetCodeNum = True
    frmFAEditItemMess.cmdExit.Text = "ESC &Return and Edit"
    frmFAEditItemMess.cmdCont.Text = "F10 &Open Department List"
    frmFAEditItemMess.Label1.Top = 900
    frmFAEditItemMess.Label1.Caption = "The asset code number entered has no match in the saved list of asset code numbers. Asset code numbers are used in the program to track and identify fixed assets. To edit the asset code number press ESC. To open a complete list of asset code numbers from which to select a valid number press F10."
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      Call cmdAssetList_Click
    Else
      Unload frmFAEditItemMess
      vaTabPro1.ActiveTab = 0
      fptxtGroupCode.SetFocus
    End If
  End If
  
  Close ACHandle
ZeroText:
   
  Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "CheckForValidAssetCodeNum", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close
End Sub

Private Sub CheckForValidTAGNum()
   Dim TagRec As FAItemRecType
   Dim THandle As Integer
   Dim TotalAccts As Integer
   Dim x As Integer
   Dim ThisText$
   
   On Error GoTo ERRORSTUFF
   
   ThisText$ = QPTrim$(fptxtTagNumber)
   If ThisText$ = "" Then Exit Sub
   
   BadTagNum = False
   
   OpenFAItemFile THandle
   TotalAccts = LOF(THandle) \ Len(TagRec)
   If TotalAccts = 0 Then
     Close
     Exit Sub
   End If
   'go thru each number one at a time and compare against all tag numbers
   For x = 1 To TotalAccts
     Get THandle, x, TagRec
       If ThisText = QPTrim$(TagRec.ItemTag) Then
         frmFAEditItemMess.Label1.Top = 900
         frmFAEditItemMess.Label1.Caption = "The tag number entered has already been assigned to another fixed asset. A unique tag number is critical in tracking fixed assets throughout the program. Please assign a unique tag number to this fixed asset. If you wish to look at a complete list of tag numbers then press F10. Otherwise press ESC to return to the screen."
         BadTagNum = True
         frmFAEditItemMess.cmdCont.Text = "F10 &Open Tag List"
         frmFAEditItemMess.cmdExit.Text = "ESC &Return and Edit"
         frmFAEditItemMess.Show vbModal
         If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
           Unload frmFAEditItemMess
           Call cmdTagList_Click
         Else
           Unload frmFAEditItemMess
           Me.vaTabPro1.ActiveTab = 0
           fptxtTagNumber.SetFocus
         End If
         Exit For
       End If
   Next x
   Close THandle
    
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "CheckForValidTAGNum", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close
End Sub

Private Function Check4ValidDept() As Boolean
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim x As Integer
  Dim NumOfDepts As Integer
  Dim CompareThis As Integer
  
  On Error GoTo ERRORSTUFF
  Check4ValidDept = True 'so far the number is fine
  
  OpenFADeptCodeFile DHandle
  NumOfDepts = LOF(DHandle) \ Len(DeptRec)
  
  If NumOfDepts = 0 Then 'no depts saved
    Close
    frmFAEditItemMess.cmdExit.Text = "ESC &Continue Saving"
    frmFAEditItemMess.cmdCont.Text = "F10 &Jump to Department Edit"
    frmFAEditItemMess.Label1.Caption = "Fixed assets need to be categorized according to the department each fixed asset is assigned. Since no departments have been set up this important fixed asset tracking feature is not possible. It is recommended that all departments are set up before you continue saving fixed assets. If you wish to jump to the department edit screen then press F10. Otherwise press ESC to return to the screen."
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      frmFAEditDeptCodes.Show
      DoEvents
      Unload Me
      Unload frmFAItemLookUp
      Check4ValidDept = False 'user warned and he elected to not save this data
    Else 'user warned and he elected to save anyway and it was recorded in the log
      Unload frmFAEditItemMess
      MainLog ("User warned that no department numbers have been saved. The user elected to save item data anyway for tag number " + fptxtTagNumber.Text + " in frmFAEditItemWTabs.")
    End If
    Exit Function 'no departments saved so no need to try to match anything
  End If
  
  CompareThis = Val(fptxtDeptNum.Text)
  For x = 1 To NumOfDepts
    Get DHandle, x, DeptRec 'start looking for a department match
    If CompareThis = Val(DeptRec.DeptNum) Then
      Exit For 'found it so we're finished
    End If
  Next x
  
  If x = NumOfDepts + 1 Then 'been thru all depts and found nothing to match
  'waht the user entered
'    MsgBox "The department number entered is not valid. Check the department list for valid department numbers."
    Check4ValidDept = False
    frmFAEditItemMess.cmdExit.Text = "ESC &Return and Edit"
    frmFAEditItemMess.cmdCont.Text = "F10 &Open Department List"
    frmFAEditItemMess.Label1.Top = 900
    frmFAEditItemMess.Label1.Caption = "The department number entered has no match in the saved list of department numbers. Department numbers are used extensively throughout the program to track and identify fixed assets. To edit the department number press ESC. To open a complete list of department numbers from which to select a valid number press F10."
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      Call cmdDept_Click
    Else
      Unload frmFAEditItemMess
      vaTabPro1.ActiveTab = 0
      fptxtDeptNum.SetFocus
    End If
  End If
  Close
  Exit Function
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "Check4ValidDept", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close
  
End Function

Private Function Check4ValidFund() As Boolean
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim x As Integer
  Dim NumOfFunds As Integer
  Dim CompareThis As Integer
  
  On Error GoTo ERRORSTUFF
  Check4ValidFund = True
  
  OpenFAFundCodeFile FHandle
  NumOfFunds = LOF(FHandle) \ Len(FundRec)
  If NumOfFunds = 0 Then
    Close
    frmFAEditItemMess.cmdExit.Text = "ESC &Continue Saving"
    frmFAEditItemMess.cmdCont.Text = "F10 &Jump to Fund Edit"
    frmFAEditItemMess.Label1.Caption = "All Fixed assets should be assigned a specific fund number. This fund number is used to keep track of the fixed asset in other areas of this program. Please make sure that all fund numbers are set up before saving any fixed assets. If you wish to jump to the fund edit screen then press F10. Otherwise press ESC to continue saving."
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      frmFAEditFundCodes.Show
      DoEvents
      Unload Me
      Unload frmFAItemLookUp
      Check4ValidFund = False 'user warned and he elected to not save this data
      Exit Function
    Else 'user warned and he elected to save anyway and it was recorded in the log
      Unload frmFAEditItemMess
      MainLog ("User warned that no fund numbers have been saved. The user elected to save item data anyway for tag number " + fptxtTagNumber.Text + " in frmFAEditItemWTabs.")
    End If
    Exit Function 'no departments saved so no need to try to match anything
  End If
  
  CompareThis = Val(fptxtFundNum.Text)
  
  For x = 1 To NumOfFunds
    Get FHandle, x, FundRec
    If CompareThis = Val(FundRec.FundNum) Then
      Exit For
    End If
  Next x
  
  If x = NumOfFunds + 1 Then
    Check4ValidFund = False
    frmFAEditItemMess.cmdExit.Text = "ESC &Return and Edit"
    frmFAEditItemMess.cmdCont.Text = "F10 &Open Fund Code List"
    frmFAEditItemMess.Label1.Top = 900
    frmFAEditItemMess.Label1.Caption = "The fund code number entered has no match in the saved list of fund code numbers. Fund code numbers are used in the program to track and identify fixed assets. To edit the fund code number press ESC. To open a complete list of fund code numbers from which to select a valid number press F10."
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "continue" Then
      Unload frmFAEditItemMess
      Call cmdFundList_Click
    Else
      Unload frmFAEditItemMess
      vaTabPro1.ActiveTab = 0
      fptxtFundNum.SetFocus
    End If
  End If
  Close
  Exit Function
  
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditItemWTabs", "Check4ValidFund", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close
  
End Function

Private Sub fptxtLeft_LostFocus()
  On Error Resume Next
  'This routine tries to protect the integrity of the life left value
  If QPTrim$(fptxtLeft.Text) = "" Then
    fptxtLeft.Text = AssLifeLeft
  ElseIf CInt(fptxtLeft.Text) <> AssLifeLeft Then 'AssLifeLeft should be the correct value
  'figured by the program...if the user changes this number to an inaccurate number and
  'then changes it again to the old correct value then the program will alert him again as
  'if the current change may be wrong
    frmFAEditItemMess.cmdExit.Text = "ESC &Return and Edit"
    frmFAEditItemMess.Label1.Height = 1500
    frmFAEditItemMess.Label1.Top = 1200
    frmFAEditItemMess.Label1.Caption = "The asset life left has been edited and may not be accurate. If you are NOT sure of the accuracy of this change then press ESC. If you want to continue with this value anyway then press F10."
    frmFAEditItemMess.Show vbModal
    If frmFAEditItemMess.fptxtChoice.Text = "abort" Then
      Unload frmFAEditItemMess
      fptxtLeft.Text = CStr(AssLifeLeft)
      fptxtLeft.SetFocus
    Else 'record this warning
      Unload frmFAEditItemMess
      MainLog ("The user changed the life left value for this asset. Item Tag # = " + TempItemTag$ + ". A warning was issued stating that the new asset life left (" + fptxtLeft.Text + " years) may not be accurate. The current asset life is " + CStr(AssLifeLeft) + " years. The user elected to save anyway in frmFAEditItemWTabs.")
    End If
  End If

End Sub

Private Sub fptxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtTagNumber.SetFocus
  End If
End Sub

Private Sub fptxtOriginalCost_LostFocus()
  Dim ORCost As Double
  Dim DepToDate As Double
  
  'update the current value based on a change in the
  'original cost
  ORCost = fptxtOriginalCost
  DepToDate = fptxtDep2Date
  fptxtCurrVal = ORCost - DepToDate

End Sub

Private Sub LogSaves()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  
  'save to the log any kind of change saved
  OpenFAItemFile FAHandle
  Get FAHandle, GRecNum, FAItemRec
  Close FAHandle
  
  If QPTrim$(TempItemTag$) <> QPTrim$(FAItemRec.ItemTag) Then
    MainLog ("Item Tag Number, " + QPTrim$(TempItemTag$) + ", changed and saved as " + QPTrim$(FAItemRec.ItemTag) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempISTATUS$) <> QPTrim$(FAItemRec.ISTATUS) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item Status, " + QPTrim$(TempISTATUS$) + ", changed and saved as " + QPTrim$(FAItemRec.ISTATUS) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempDEPYN$) <> QPTrim$(FAItemRec.DEPYN) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Depreciate Y/N?, " + QPTrim$(TempDEPYN$) + ", changed and saved as " + QPTrim$(FAItemRec.DEPYN) + " in frmFAEditItemWTabs.")
  End If
  
  If TempAQURDATE <> FAItemRec.AQURDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item acquire date, " + MakeRegDate(TempAQURDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.AQURDATE) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempIDESC1$) <> QPTrim$(FAItemRec.IDESC1) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item description 1,  " + QPTrim$(TempIDESC1$) + ", changed and saved as " + QPTrim$(FAItemRec.IDESC1) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempIDESC2$) <> QPTrim$(FAItemRec.IDESC2) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item description 2, " + QPTrim$(TempIDESC2$) + ", changed and saved as " + QPTrim$(FAItemRec.IDESC2) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempGLACCT$) <> QPTrim$(FAItemRec.GLACCT) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item GL acct number, " + QPTrim$(TempGLACCT$) + ", changed and saved as " + QPTrim$(FAItemRec.GLACCT) + " in frmFAEditItemWTabs.")
  End If
  
  If TempIDEPT <> FAItemRec.IDEPT Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item department number, " + CStr(TempIDEPT) + ", changed and saved as " + CStr(FAItemRec.IDEPT) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempASSETCode$) <> QPTrim$(FAItemRec.ASSETCODE) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item asset code number, " + QPTrim$(TempASSETCode$) + ", changed and saved as " + QPTrim$(FAItemRec.ASSETCODE) + " in frmFAEditItemWTabs.")
  End If
  
  If TempILIFE <> FAItemRec.ILIFE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item life, " + CStr(TempILIFE) + ", changed and saved as " + CStr(FAItemRec.ILIFE) + " in frmFAEditItemWTabs.")
  End If
  
  If TempORGCOST <> FAItemRec.ORGCOST Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item purchase price, " + CStr(TempORGCOST) + ", changed and saved as " + CStr(FAItemRec.ORGCOST) + " in frmFAEditItemWTabs.")
  End If
  
  If TempDEP2DATE <> FAItemRec.DEP2DATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item depreciation to date,  " + CStr(TempDEP2DATE) + ", changed and saved as " + CStr(FAItemRec.DEP2DATE) + " in frmFAEditItemWTabs.")
  End If
  
  If TempCURRVAL <> FAItemRec.CURRVAL Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item current value, " + CStr(TempCURRVAL) + ", changed and saved as " + CStr(FAItemRec.CURRVAL) + " in frmFAEditItemWTabs.")
  End If
  
  If FAItemRec.CDEPDATE < -11000 Then FAItemRec.CDEPDATE = 0
  If TempCDEPDATE <> FAItemRec.CDEPDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item last depreciation date, " + MakeRegDate(TempCDEPDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.CDEPDATE) + " in frmFAEditItemWTabs.")
  End If
  
  If TempDispDate <> FAItemRec.DispDate Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item disposal date, " + MakeRegDate(TempDispDate) + ", changed and saved as " + MakeRegDate(FAItemRec.DispDate) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVENDOR$) <> QPTrim$(FAItemRec.VENDOR) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vendor, " + QPTrim$(TempVENDOR$) + ", changed and saved as " + QPTrim$(FAItemRec.VENDOR) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempSERIALNO$) <> QPTrim$(FAItemRec.SERIALNO) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item serial number, " + QPTrim$(TempSERIALNO$) + ", changed and saved as " + QPTrim$(FAItemRec.SERIALNO) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempITEMMFG$) <> QPTrim$(FAItemRec.ITEMMFG) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item manufacturer, " + QPTrim$(TempITEMMFG$) + ", changed and saved as " + QPTrim$(FAItemRec.ITEMMFG) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempCONTACT$) <> QPTrim$(FAItemRec.CONTACT) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item contact, " + QPTrim$(TempCONTACT$) + ", changed and saved as " + QPTrim$(FAItemRec.CONTACT) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempPhone$) <> QPTrim$(FAItemRec.Phone) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item phone, " + QPTrim$(TempPhone$) + ", changed and saved as " + QPTrim$(FAItemRec.Phone) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempITEMLOC$) <> QPTrim$(FAItemRec.ITEMLOC) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item location, " + QPTrim$(TempITEMLOC$) + ", changed and saved as " + QPTrim$(FAItemRec.ITEMLOC) + " in frmFAEditItemWTabs.")
  End If
  
  If TempEOLDATE <> FAItemRec.EOLDATE Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item end of life, " + MakeRegDate(TempEOLDATE) + ", changed and saved as " + MakeRegDate(FAItemRec.EOLDATE) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVHCLMAKE$) <> QPTrim$(FAItemRec.VHCLMAKE) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle make, " + QPTrim$(TempVHCLMAKE$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLMAKE) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVHCLMODL$) <> QPTrim$(FAItemRec.VHCLMODL) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle model, " + QPTrim$(TempVHCLMODL$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLMODL) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVHCLVIN$) <> QPTrim$(FAItemRec.VHCLVIN) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle ID number, " + QPTrim$(TempVHCLVIN$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLVIN) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVHCLTAG$) <> QPTrim$(FAItemRec.VHCLTAG) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle license tag number, " + QPTrim$(TempVHCLTAG$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLTAG) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempVHCLCOLR$) <> QPTrim$(FAItemRec.VHCLCOLR) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item vehicle color," + QPTrim$(TempVHCLCOLR$) + ", changed and saved as " + QPTrim$(FAItemRec.VHCLCOLR) + " in frmFAEditItemWTabs.")
  End If
  
  If TempWARRXDAT <> FAItemRec.WARRXDAT Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item warranty expiration date, " + MakeRegDate(TempWARRXDAT) + ", changed and saved as " + MakeRegDate(FAItemRec.WARRXDAT) + " in frmFAEditItemWTabs.")
  End If
  
  If TempFundNum <> FAItemRec.FundNum Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item fund number, " + CStr(TempFundNum) + ", changed and saved as " + CStr(FAItemRec.FundNum) + " in frmFAEditItemWTabs.")
  End If
  
  If TempDisposAmt <> FAItemRec.DisposAmt Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item disposal amount, " + CStr(TempDisposAmt) + ", changed and saved as " + CStr(FAItemRec.DisposAmt) + " in frmFAEditItemWTabs.")
  End If
  
  If TempLifeLeft <> FAItemRec.LifeLeft Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item life left, " + CStr(TempLifeLeft) + ", changed and saved as " + CStr(FAItemRec.LifeLeft) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempPONum$) <> QPTrim$(FAItemRec.PONum) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item purchase order number, " + QPTrim$(TempPONum$) + ", changed and saved as " + QPTrim$(FAItemRec.PONum) + " in frmFAEditItemWTabs.")
  End If
  
  If QPTrim$(TempCheckNum$) <> QPTrim$(FAItemRec.CheckNum) Then
    MainLog ("For tag number " + QPTrim$(FAItemRec.ItemTag) + ":" + "Item check number, " + QPTrim$(TempCheckNum$) + ", changed and saved as " + QPTrim$(FAItemRec.CheckNum) + " in frmFAEditItemWTabs.")
  End If

End Sub

Private Sub fptxtTagNumber_LostFocus()
  fptxtHeader.Text = QPTrim$(fptxtTagNumber.Text) + "  " + QPTrim$(fptxtDesc1.Text)
End Sub

Private Sub fptxtVhclColr_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtVhclMake.SetFocus
  End If
End Sub

