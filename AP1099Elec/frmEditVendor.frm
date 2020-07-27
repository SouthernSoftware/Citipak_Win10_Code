VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditVendor 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit A Vendor"
   ClientHeight    =   8640
   ClientLeft      =   150
   ClientTop       =   -2175
   ClientWidth     =   12225
   Icon            =   "frmEditVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboActive 
      Height          =   375
      Left            =   7770
      TabIndex        =   61
      Top             =   2010
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
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
      BackColor       =   16777215
      ForeColor       =   0
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
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmEditVendor.frx":08CA
   End
   Begin LpLib.fpList fplstVendors 
      Height          =   1740
      Left            =   405
      TabIndex        =   56
      Top             =   5325
      Visible         =   0   'False
      Width           =   5925
      _Version        =   196608
      _ExtentX        =   10451
      _ExtentY        =   3069
      TextAlias       =   ""
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
      BackColor       =   16777215
      ForeColor       =   0
      Columns         =   3
      Sorted          =   1
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   8421504
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   12632256
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
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmEditVendor.frx":0CFF
   End
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   4140
      Left            =   1728
      TabIndex        =   28
      Top             =   2472
      Width           =   8772
      _Version        =   196609
      _ExtentX        =   15473
      _ExtentY        =   7302
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
      ForeColor       =   0
      TabCount        =   3
      ThreeD          =   -1  'True
      ThreeDShadowColor=   4210752
      ThreeDHighlightColor=   14737632
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ActiveTabBold   =   -1  'True
      GrayAreaColor   =   9405029
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      EarMarkColorDark=   13684944
      EarMarkColorLight=   14737632
      EarMarkColorOutline=   12632256
      BookRingColorLight=   12632256
      BookShowMetalSpine=   -1  'True
      BookRingShowHole=   -1  'True
      PageEarMarkColorDark=   12632256
      PageEarMarkColorLight=   14737632
      PageEarMarkColorOutline=   13684944
      DataFormat      =   ""
      AutoSizeChildren=   3
      BookCornerGuardWidth=   90
      BookCornerGuardLength=   375
      BookCornerGuardColor=   32896
      ThreeDOuterDark =   0
      ThreeDOuterLight=   16777215
      ThreeDInnerDark =   8421504
      ThreeDInnerLight=   14737632
      EarMarkPictureMaskColor=   14737632
      DataField       =   ""
      TabCaption      =   "frmEditVendor.frx":112F
      PageEarMarkPictureNext=   "frmEditVendor.frx":142B
      PageEarMarkPicturePrev=   "frmEditVendor.frx":1447
      EarMarkPictureNext=   "frmEditVendor.frx":1463
      EarMarkPicturePrev=   "frmEditVendor.frx":147F
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   3132
         Left            =   -20184
         TabIndex        =   29
         Top             =   -15480
         Width           =   8148
         _Version        =   196609
         _ExtentX        =   14372
         _ExtentY        =   5524
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   13684944
         Caption         =   ""
         ForeColor       =   0
         FrameColor      =   12632256
         FrameThreeDHighlightColor=   14737632
         FrameThreeDShadowColor=   8421504
         FrameThreeDStyle=   2
         OutlineInsideColor=   4210752
         OutlineOutsideColor=   0
         ThreeDHighlightColor=   14737632
         ThreeDShadowColor=   8421504
         ThreeDStyle     =   0
         ThreeDTextHighlightColor=   14737632
         ThreeDTextShadowColor=   8421504
         Picture         =   "frmEditVendor.frx":149B
         Begin LpLib.fpCombo fpcboVGet1099 
            Height          =   405
            Left            =   1890
            TabIndex        =   18
            Top             =   2430
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmEditVendor.frx":14B7
         End
         Begin EditLib.fpMask fpmskVFax 
            Height          =   396
            Left            =   4680
            TabIndex        =   21
            Top             =   1344
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
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   0   'False
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
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "\1(###)###-####"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   -1  'True
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
         Begin EditLib.fpMask fpmskVPhone 
            Height          =   396
            Left            =   4680
            TabIndex        =   20
            Top             =   792
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
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "\1(###)###-####"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   -1  'True
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
         Begin EditLib.fpText fptxtVContact 
            Height          =   396
            Left            =   4680
            TabIndex        =   19
            Top             =   216
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
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVCountyCode 
            Height          =   396
            Left            =   1896
            TabIndex        =   17
            Top             =   1884
            Width           =   636
            _Version        =   196608
            _ExtentX        =   1122
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   3
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
         Begin EditLib.fpText fptxtVStateCode 
            Height          =   396
            Left            =   1896
            TabIndex        =   16
            Top             =   1344
            Width           =   636
            _Version        =   196608
            _ExtentX        =   1122
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVFedId 
            Height          =   396
            Left            =   1896
            TabIndex        =   15
            Top             =   804
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVTerms 
            Height          =   396
            Left            =   1896
            TabIndex        =   14
            Top             =   264
            Width           =   708
            _Version        =   196608
            _ExtentX        =   1249
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "0123456789"
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
         Begin EditLib.fpText fptxtDBA 
            Height          =   396
            Left            =   3864
            TabIndex        =   22
            Top             =   2496
            Width           =   4140
            _Version        =   196608
            _ExtentX        =   7302
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0D0D0&
            Caption         =   "DBA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3120
            TabIndex        =   58
            Top             =   2592
            Width           =   636
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Phone/Fax Format Example: 1(987)111-2222"
            Height          =   492
            Left            =   4536
            TabIndex        =   57
            Top             =   1848
            Width           =   2460
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   3912
            TabIndex        =   37
            Top             =   1392
            Width           =   540
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   3672
            TabIndex        =   36
            Top             =   876
            Width           =   804
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   3528
            TabIndex        =   35
            Top             =   312
            Width           =   972
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Get 1099"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   624
            TabIndex        =   34
            Top             =   2472
            Width           =   1068
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "County Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   276
            Left            =   744
            TabIndex        =   33
            Top             =   1992
            Width           =   948
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "State Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   324
            Left            =   480
            TabIndex        =   32
            Top             =   1488
            Width           =   1212
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Federal Id#"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   348
            Left            =   408
            TabIndex        =   31
            Top             =   888
            Width           =   1284
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   348
            Left            =   912
            TabIndex        =   30
            Top             =   336
            Width           =   780
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   3132
         Left            =   -20184
         TabIndex        =   46
         Top             =   -15480
         Width           =   8148
         _Version        =   196609
         _ExtentX        =   14372
         _ExtentY        =   5524
         _StockProps     =   70
         Enabled         =   0   'False
         BackColor       =   14737632
         Caption         =   ""
         ForeColor       =   0
         FrameColor      =   12632256
         FrameThreeDHighlightColor=   14737632
         FrameThreeDShadowColor=   8421504
         FrameThreeDStyle=   2
         ThreeDHighlightColor=   14737632
         ThreeDShadowColor=   8421504
         ThreeDStyle     =   0
         ThreeDTextHighlightColor=   14737632
         ThreeDTextShadowColor=   8421504
         Picture         =   "frmEditVendor.frx":18B5
         Begin VB.CommandButton cmdCopy 
            BackColor       =   &H00D0D0D0&
            Caption         =   "F4 Cop&y Vendor Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1068
            Left            =   6648
            MaskColor       =   &H00D0D0D0&
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   408
            Width           =   996
         End
         Begin EditLib.fpMask fpmskChkZip 
            Height          =   396
            Left            =   6408
            TabIndex        =   12
            Top             =   1896
            Width           =   1380
            _Version        =   196608
            _ExtentX        =   2434
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483634
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483634
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "AAAAA-AAAA"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   0   'False
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
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
         Begin EditLib.fpText fptxtChkState 
            Height          =   396
            Left            =   5160
            TabIndex        =   11
            Top             =   1896
            Width           =   564
            _Version        =   196608
            _ExtentX        =   995
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtChkCity 
            Height          =   396
            Left            =   1728
            TabIndex        =   10
            Top             =   1899
            Width           =   2556
            _Version        =   196608
            _ExtentX        =   4508
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   22
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
         Begin EditLib.fpText fptxtChkName 
            Height          =   396
            Left            =   1728
            TabIndex        =   7
            Top             =   324
            Width           =   4236
            _Version        =   196608
            _ExtentX        =   7472
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtChkAdd1 
            Height          =   396
            Left            =   1728
            TabIndex        =   8
            Top             =   849
            Width           =   4260
            _Version        =   196608
            _ExtentX        =   7514
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtChkAdd2 
            Height          =   396
            Left            =   1728
            TabIndex        =   9
            Top             =   1374
            Width           =   4260
            _Version        =   196608
            _ExtentX        =   7514
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtChkMemo 
            Height          =   396
            Left            =   1728
            TabIndex        =   13
            Top             =   2424
            Width           =   4140
            _Version        =   196608
            _ExtentX        =   7302
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Check Memo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   48
            TabIndex        =   60
            Top             =   2472
            Width           =   1572
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   744
            TabIndex        =   52
            Top             =   420
            Width           =   852
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   696
            TabIndex        =   51
            Top             =   924
            Width           =   924
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   396
            Left            =   504
            TabIndex        =   50
            Top             =   1440
            Width           =   1116
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1008
            TabIndex        =   49
            Top             =   1944
            Width           =   564
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
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
            Height          =   348
            Left            =   4440
            TabIndex        =   48
            Top             =   1920
            Width           =   612
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   5832
            TabIndex        =   47
            Top             =   1920
            Width           =   468
         End
      End
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   3180
         Left            =   315
         TabIndex        =   38
         Top             =   615
         Width           =   8130
         _Version        =   196609
         _ExtentX        =   14340
         _ExtentY        =   5609
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
         BackColor       =   13684944
         Caption         =   ""
         FrameColor      =   12632256
         FrameThreeDHighlightColor=   14737632
         FrameThreeDShadowColor=   8421504
         FrameThreeDStyle=   2
         OutlineInsideColor=   0
         OutlineOutsideColor=   0
         ThreeDHighlightColor=   14737632
         ThreeDShadowColor=   8421504
         ThreeDStyle     =   0
         ThreeDTextHighlightColor=   14737632
         ThreeDTextShadowColor=   8421504
         Picture         =   "frmEditVendor.frx":18D1
         Begin LpLib.fpCombo fpcboVendCode 
            Height          =   405
            Left            =   1770
            TabIndex        =   0
            Top             =   270
            Width           =   6045
            _Version        =   196608
            _ExtentX        =   10663
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
            ColumnSearch    =   0
            ColumnWidthScale=   2
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   1
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
            ScrollBarH      =   3
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
            AutoSearchFillDelay=   100
            EditMarginLeft  =   5
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmEditVendor.frx":18ED
         End
         Begin EditLib.fpText fptxtVendCode 
            Height          =   396
            Left            =   4080
            TabIndex        =   39
            Top             =   -1656
            Width           =   2436
            _Version        =   196608
            _ExtentX        =   4297
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
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "0123456789"
            MaxLength       =   10
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
         Begin EditLib.fpMask fpmskVendZip 
            Height          =   396
            Left            =   6456
            TabIndex        =   6
            Top             =   2424
            Width           =   1380
            _Version        =   196608
            _ExtentX        =   2434
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
            BackColor       =   16777215
            ForeColor       =   0
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
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "AAAAA-AAAA"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   0   'False
            BorderGrayAreaColor=   12632256
            NoPrefix        =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   8421504
            BorderDropShadowWidth=   3
            AutoTab         =   0   'False
            ButtonColor     =   12632256
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fptxtVendState 
            Height          =   396
            Left            =   5208
            TabIndex        =   5
            Top             =   2424
            Width           =   516
            _Version        =   196608
            _ExtentX        =   910
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVendCity 
            Height          =   396
            Left            =   1776
            TabIndex        =   4
            Top             =   2424
            Width           =   2556
            _Version        =   196608
            _ExtentX        =   4508
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   22
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
         Begin EditLib.fpText fptxtVendAdd1 
            Height          =   396
            Left            =   1776
            TabIndex        =   2
            Top             =   1344
            Width           =   4140
            _Version        =   196608
            _ExtentX        =   7302
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVendAdd2 
            Height          =   396
            Left            =   1776
            TabIndex        =   3
            Top             =   1884
            Width           =   4140
            _Version        =   196608
            _ExtentX        =   7302
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin EditLib.fpText fptxtVendName 
            Height          =   396
            Left            =   1776
            TabIndex        =   1
            Top             =   816
            Width           =   4140
            _Version        =   196608
            _ExtentX        =   7302
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   2
            CaretOverWrite  =   2
            UserEntry       =   1
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   2
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   72
            TabIndex        =   53
            Top             =   840
            Width           =   1596
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   5880
            TabIndex        =   45
            Top             =   2448
            Width           =   468
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Height          =   348
            Left            =   4488
            TabIndex        =   44
            Top             =   2448
            Width           =   612
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1056
            TabIndex        =   43
            Top             =   2424
            Width           =   564
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Height          =   396
            Left            =   552
            TabIndex        =   42
            Top             =   1824
            Width           =   1116
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   744
            TabIndex        =   41
            Top             =   1236
            Width           =   924
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   240
            TabIndex        =   40
            Top             =   336
            Width           =   1428
         End
      End
   End
   Begin EditLib.fpText fpViewVen 
      Height          =   348
      Left            =   2616
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2016
      Visible         =   0   'False
      Width           =   5148
      _Version        =   196608
      _ExtentX        =   9080
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
      BackColor       =   12632256
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
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
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
      Appearance      =   3
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   8280
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:31 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2/12/2008"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   468
      Left            =   6696
      TabIndex        =   23
      Top             =   7272
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmEditVendor.frx":1D6F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVList 
      Height          =   468
      Left            =   5208
      TabIndex        =   55
      Top             =   7272
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmEditVendor.frx":1F4F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   468
      Left            =   9696
      TabIndex        =   24
      Top             =   7248
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   825
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
      ButtonDesigner  =   "frmEditVendor.frx":212E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   465
      Left            =   8190
      TabIndex        =   25
      Top             =   7245
      Width           =   1320
      _Version        =   131072
      _ExtentX        =   2328
      _ExtentY        =   820
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
      ButtonDesigner  =   "frmEditVendor.frx":230E
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      FillColor       =   &H80000009&
      Height          =   4308
      Left            =   1638
      Top             =   2376
      Width           =   8940
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit A Vendor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4884
      TabIndex        =   27
      Top             =   936
      Width           =   2460
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2592
      Top             =   672
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   552
      Width           =   7020
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
Attribute VB_Name = "frmEditVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Vendor As VendorRecType
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Dim NumOpenItems As Integer
Dim allowus2edit As Boolean
Private Sub cmdExit_Click()
  If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
    If Changed Then
      If MsgBox("Exit And Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
        Exit Sub
      End If
    End If
  End If
  frmAPVendMaintMenu.Show
  Unload frmEditVendor
End Sub
Private Sub cmdCopy_Click()
  fptxtChkName.Text = fptxtVendName.Text
  fptxtChkAdd1.Text = fptxtVendAdd1.Text
  fptxtChkAdd2.Text = fptxtVendAdd2.Text
  fptxtChkCity.Text = fptxtVendCity.Text
  fptxtChkState.Text = fptxtVendState.Text
  fpmskChkZip.Text = fpmskVendZip.Text
  
End Sub

Private Sub cmdVList_Click()
  fplstVendors.Visible = True
  fplstVendors.ZOrder (0)
  fplstVendors.SetFocus
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ((UnloadMode = vbFormControlMenu)) Then
  If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbYes Then
    If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
      If Changed Then
        If MsgBox("Close And Abandon Changes?", vbYesNo, "Close?") = vbNo Then
          Cancel = True
        Else
          MainLog "Close AP"
          ClearInUse PWcnt
        End If
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      MainLog "Close AP"
      ClearInUse PWcnt
    End If
  Else
    Cancel = True
  End If
End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEditVend
  vaTabPro1.ActiveTab = 0
  VendCodeNameIA fpcboVendCode
  fpcboVGet1099.AddItem "Yes"
  fpcboVGet1099.AddItem "No"
  fptxtVTerms = 0
  VendsLstAlpha fplstVendors
  fpcboActive.AddItem "Active"
  fpcboActive.AddItem "InActive"
  fpcboActive.ListIndex = -1
  If PWUser$ = "Sosoft Support" Then
    allowus2edit = True
  Else
    allowus2edit = False
  End If
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
 '   Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case vbKeyF5:
      cmdVList_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case vbKeyPageUp:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 2
        fptxtVTerms.SetFocus
      Else
        If vaTabPro1.ActiveTab = 1 Then
          vaTabPro1.ActiveTab = 0
          fpcboVendCode.SetFocus
        Else
          If vaTabPro1.ActiveTab = 2 Then
            vaTabPro1.ActiveTab = 1
            fptxtChkName.SetFocus
          End If
        End If
      End If
      KeyCode = 0
    Case vbKeyPageDown:
      If vaTabPro1.ActiveTab = 0 Then
        vaTabPro1.ActiveTab = 1
        fptxtChkName.SetFocus
      Else
        If vaTabPro1.ActiveTab = 1 Then
          vaTabPro1.ActiveTab = 2
          fptxtVTerms.SetFocus
        Else
          If vaTabPro1.ActiveTab = 2 Then
            vaTabPro1.ActiveTab = 0
            fpcboVendCode.SetFocus
          End If
        End If
      End If
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub fpcboVendCode_BeforeDropDown(Cancel As Boolean)
    If Changed Then
      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
        Cancel = True
      Else
       ' clearscreen
        fpcboVendCode.ListIndex = fpcboVendCode.ListIndex - 1
      End If
    End If
End Sub

Private Sub fpcboVendCode_Change()
'  If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
'
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        'Cancel = True
'        fptxtVendName.SetFocus
'      Else
'        LoadVendor
'      End If
'    Else
      LoadVendor
'    End If
'   End If
End Sub

'Private Sub fpcboVendCode_DropDown()
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        'fpcboVendCode.ListDown = False
'        fpcboVendCode.Action = ActionDeselectAll
'      Else
'       ' clearscreen
'        fpcboVendCode
'        'fpcboVendCode.ListDown = True
'      End If
'    End If
'    fpcboVendCode.Action = ActionClearSearchBuffer
'End Sub

'Private Sub fpcboVendCode_BeforeDropDown(Cancel As Boolean)
'   If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        Cancel = True
'        fptxtVendName.SetFocus
'      End If
'    End If
'   End
'
'End Sub

'Private Sub fpcboVendCode_Change()
'    If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        'Cancel = True
'        fptxtVendName.SetFocus
'      End If
'    End If
'   End If
'
'End Sub
'Private Sub fpcboVendCode_KeyPress(KeyAscii As Integer)
'  If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'
'        fptxtVendName.SetFocus
'      End If
'    End If
'  End If
'
'End Sub

'Private Sub fpcboVendCode_CloseUp()
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        'Cancel = True
'        fptxtVendName.SetFocus
'      Else
'        LoadVendor
'      End If
'    Else
'      LoadVendor
'    End If
'End Sub
'
'Private Sub fpcboVendCode_KeyUp(KeyCode As Integer, Shift As Integer)
' Stop
'End Sub
'
'Private Sub fpcboVendCode_SelChange(ItemIndex As Long)
'    If Changed Then
'      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
'        'Cancel = True
'        fptxtVendName.SetFocus
'      Else
'        LoadVendor
'      End If
'    Else
'      LoadVendor
'    End If
'End Sub

Private Sub fplstVendors_LostFocus()
  fplstVendors.Visible = False
End Sub
'Private Sub fptxtVendCode_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyUp, vbKeyLeft:
'      SendKeys "{pgup}"
'    Case vbKeyDown, vbKeyRight:
'      fptxtVendName.SetFocus
'      KeyCode = 0
'    Case Else:
'  End Select
'
'End Sub
Private Sub fpcboVendCode_KeyDown(KeyCode As Integer, Shift As Integer)
 On Local Error Resume Next
  If KeyCode = vbKeySpace Then
    fpcboVendCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVendCode.ListIndex = -1
    fpcboVendCode.Action = ActionClearSearchBuffer
  End If
  If fpcboVendCode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtVendName.SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        'SendKeys "{pgup}"
        vaTabPro1.ActiveTab = 2
        fptxtVTerms.SetFocus
        KeyCode = 0
    Else
      fpcboVendCode.ListDown = True
    End If
  End If

End Sub

Private Sub fpmskChkZip_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      fptxtChkMemo.SetFocus
    Case vbKeyUp, vbKeyLeft:
      fptxtChkState.SetFocus
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fptxtChkMemo_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      SendKeys "{pgdn}"
    Case vbKeyUp, vbKeyLeft:
      fpmskChkZip.SetFocus
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub fpmskVendZip_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      SendKeys "{pgdn}"
    Case vbKeyUp, vbKeyLeft:
      fptxtVendState.SetFocus
      KeyCode = 0
    Case Else:
   End Select
End Sub

Private Sub fptxtVTerms_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyLeft:
      SendKeys "{pgup}"
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      fptxtVFedId.SetFocus
      KeyCode = 0
    Case Else:
   End Select
End Sub

Private Sub fpmskVFax_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      fptxtDBA.SetFocus
      KeyCode = 0
    Case vbKeyUp, vbKeyLeft:
      fpmskVPhone.SetFocus
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fptxtDBA_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      SendKeys "{pgdn}"
    Case vbKeyUp, vbKeyLeft:
      fpmskVFax.SetFocus
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub fptxtChkName_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyLeft:
      SendKeys "{pgup}"
    Case vbKeyDown, vbKeyRight, vbKeyReturn:
      fptxtChkAdd1.SetFocus
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub fpcboActive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboActive.ListDown = True
  End If
  If fpcboActive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fplstVendors_DblClick()
  Dim vrec As String

  fpcboVendCode.SearchIndex = 0
  fplstVendors.col = 2
  vrec = QPTrim(fplstVendors.ColText)
  'Ya gotta refer to all this search stuff so will search the
  'correct column then set back to original so when key vendor code
  'will find correct one.
  fpcboVendCode.ColumnSearch = 2
  fpcboVendCode.SearchText = vrec
  fpcboVendCode.SearchMethod = SearchMethodExactMatch
  fpcboVendCode.Action = ActionSearch
  If fpcboVendCode.SearchIndex <> -1 Then
   If fpcboVendCode.ListIndex <> -1 And fptxtVendName <> "" Then
    If Changed Then
      If MsgBox("Abandon Changes?", vbYesNo, "Exit?") = vbNo Then
        fpcboVendCode.SearchIndex = 0
        fplstVendors.Visible = False
        fpcboVendCode.ColumnSearch = 0
        Exit Sub
      End If
    End If
   End If

    fpcboVendCode.ListIndex = fpcboVendCode.SearchIndex
  End If
  fpcboVendCode.SearchIndex = 0
  fplstVendors.Visible = False
  fpcboVendCode.ColumnSearch = 0
End Sub

'Private Sub fpcboVendCode_Click()
'  If fpcboVendCode.ListIndex <> -1 Then
'    LoadVendor
'  End If
'End Sub
Private Sub fpcboVGet1099_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVGet1099.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVGet1099.ListIndex = -1
    fpcboVGet1099.Action = ActionClearSearchBuffer
  End If
  If fpcboVGet1099.ListDown <> True Then
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

Private Sub fpcboVendCode_LostFocus()
  fpcboVendCode.Action = ActionClearSearchBuffer
  
End Sub
Private Sub LoadVendor()
  Dim VendorFile As Integer, NumVRecs As Integer, vnum As Integer
 On Local Error GoTo loadvenerr
10:
  ChkVenActivity
20:
  OpenVendorFile VendorFile, NumVRecs
  fpcboVendCode.col = 2
  vnum = Val(QPTrim(fpcboVendCode.ColText))
  If vnum > 0 Then
30:
  Get VendorFile, vnum, Vendor
  fpViewVen.Visible = True
  fpViewVen.Text = (Vendor.vnum + " " + Vendor.VNAME)
  If Vendor.ActiveFlag = 1 Then
    fpcboActive.ListIndex = 1
  Else
    fpcboActive.ListIndex = 0
  End If
  If NumOpenItems > 0 And Vendor.ActiveFlag = 1 And allowus2edit = False Then
    fpcboActive.Enabled = False
  Else
    fpcboActive.Enabled = True
  End If
  fptxtVendName = Vendor.VNAME
  fptxtVendAdd1 = Vendor.Addr1
  fptxtVendAdd2 = Vendor.Addr2
  fptxtVendCity = Vendor.City
  fptxtVendState = Vendor.State
  fpmskVendZip = Vendor.Zip
  fptxtChkName = Vendor.PaytoName
  fptxtChkAdd1 = Vendor.PaytoAddr
  fptxtChkAdd2 = Vendor.PaytoAddr2
  fptxtChkCity = Vendor.PayToCity
  fptxtChkState = Vendor.PaytoState
  fpmskChkZip = Vendor.PaytoZip
  fptxtChkMemo = Vendor.Memo
  fptxtVTerms = QPTrim$(Str$(Vendor.VTerms))
  fptxtVFedId = Vendor.Fedid
  fptxtVStateCode = Vendor.StCode
  fptxtVCountyCode = Vendor.CoCode
40:
  Select Case Vendor.Get1099
    Case "Y"
      fpcboVGet1099.ListIndex = 0
    Case "N"
      fpcboVGet1099.ListIndex = 1
    Case Else
      fpcboVGet1099.ListIndex = -1
  End Select
  fptxtVContact = QPTrim(Vendor.Contact)
  fpmskVPhone = QPTrim(Vendor.Phone)
  fpmskVFax = QPTrim(Vendor.Fax)
  fptxtDBA = QPTrim(Vendor.DBA)
End If
Close VendorFile
loadvenerr:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (loadven - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
Private Function Changed()
  Dim cnt As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, vrec As Integer
  OpenVendorFile VendorFile, NumVRecs
  fpcboVendCode.col = 2
  vrec = Val(fpcboVendCode.ColText)
  If vrec > 0 Then
  Get VendorFile, vrec, Vendor
  If QPTrim(fptxtVendName) <> QPTrim(Vendor.VNAME) Then cnt = cnt + 1
  If QPTrim(fptxtVendAdd1) <> QPTrim(Vendor.Addr1) Then cnt = cnt + 1
  If QPTrim(fptxtVendAdd2) <> QPTrim(Vendor.Addr2) Then cnt = cnt + 1
  If QPTrim(fptxtVendCity) <> QPTrim(Vendor.City) Then cnt = cnt + 1
  If QPTrim(fptxtVendState) <> QPTrim(Vendor.State) Then cnt = cnt + 1
  If QPTrim(fpmskVendZip) <> QPTrim(Vendor.Zip) Then cnt = cnt + 1
  If QPTrim(fptxtChkName) <> QPTrim(Vendor.PaytoName) Then cnt = cnt + 1
  If QPTrim(fptxtChkAdd1) <> QPTrim(Vendor.PaytoAddr) Then cnt = cnt + 1
  If QPTrim(fptxtChkAdd2) <> QPTrim(Vendor.PaytoAddr2) Then cnt = cnt + 1
  If QPTrim(fptxtChkCity) <> QPTrim(Vendor.PayToCity) Then cnt = cnt + 1
  If QPTrim(fptxtChkState) <> QPTrim(Vendor.PaytoState) Then cnt = cnt + 1
  If QPTrim(fpmskChkZip) <> QPTrim(Vendor.PaytoZip) Then cnt = cnt + 1
  If QPTrim(fptxtChkMemo) <> QPTrim(Vendor.Memo) Then cnt = cnt + 1
  If QPTrim(fptxtVTerms) <> QPTrim$(Str$(Vendor.VTerms)) Then cnt = cnt + 1
  If QPTrim(fptxtVFedId) <> QPTrim(Vendor.Fedid) Then cnt = cnt + 1
  If QPTrim(fptxtVStateCode) <> QPTrim(Vendor.StCode) Then cnt = cnt + 1
  If QPTrim(fptxtVCountyCode) <> QPTrim(Vendor.CoCode) Then cnt = cnt + 1
  If QPTrim(Vendor.Get1099) <> Mid$(QPTrim(fpcboVGet1099.Text), 1, 1) Then cnt = cnt + 1
  If QPTrim(fptxtVContact) <> QPTrim(Vendor.Contact) Then cnt = cnt + 1
  If QPTrim(fpmskVPhone.Text) <> "1(" Then
    If QPTrim(fpmskVPhone.Text) <> QPTrim(Vendor.Phone) Then cnt = cnt + 1
  End If
  If QPTrim(fpmskVFax.Text) <> "1(" Then
    If QPTrim(fpmskVFax.Text) <> QPTrim(Vendor.Fax) Then cnt = cnt + 1
  End If
  If QPTrim(fptxtDBA.Text) <> QPTrim(Vendor.DBA) Then cnt = cnt + 1
  If Vendor.ActiveFlag = 0 Then
    If fpcboActive.ListIndex <> 0 Then cnt = cnt + 1
  Else
    If fpcboActive.ListIndex <> 1 Then cnt = cnt + 1
  End If
  If cnt > 0 Then Changed = True
  End If
Close VendorFile

End Function
Private Sub cmdSave_Click()
  If Len(QPTrim$(fptxtVendName)) > 0 Then
      If Not Len(QPTrim$(fptxtChkName.Text)) > 0 Or Not Len(QPTrim$(fptxtChkAdd1.Text)) > 0 Or Not Len(QPTrim$(fptxtChkCity.Text)) > 0 Or Not Len(QPTrim$(fptxtChkState.Text)) > 0 Or fpmskChkZip.Text = "" Then
        GoSub DoMsg
      End If
      If chkcodes = False Then
         GoSub domsg2
      End If

    SaveVend
    MsgBox "Save Completed.", vbOKOnly, "Completed"
    clearscreen
    
  Else
    MsgBox "You May Not Leave The Vendor Name Blank.", vbOKOnly, "Denied"
    vaTabPro1.ActiveTab = 0
    fptxtVendName.SetFocus
  End If
  Exit Sub
  
domsg2:
      If MsgBox("Leaving the County and State Codes Blank will not provide the needed information to generate the Sales Tax Reports, Continue With Save, or Cancel and Edit ?", vbOKCancel, "Continue") = vbCancel Then
        vaTabPro1.ActiveTab = 2
        fptxtVStateCode.SetFocus
        Exit Sub
      End If
Return
  
DoMsg:
      If MsgBox("Check Information Not Complete, Continue With Save, or Cancel and Edit ?", vbOKCancel, "Continue") = vbCancel Then
        vaTabPro1.ActiveTab = 1
        fptxtChkName.SetFocus
        Exit Sub
      Else
        If Not Len(QPTrim$(fptxtChkName.Text)) > 0 Then fptxtChkName.Text = fptxtVendName.Text
      End If
 Return
End Sub
Private Function chkcodes()
  If fptxtVStateCode = "" Or fptxtVCountyCode = "" Then
    chkcodes = False
  Else
    chkcodes = True
  End If
End Function

Private Sub SaveVend()
  Dim VendorFile As Integer, NumVRecs As Integer, vrec As Integer
  On Local Error GoTo saveerr
  'PrintHelp "Saving to disk.. Please wait."
10:
  OpenVendorFile VendorFile, NumVRecs
  'Vendor.VNum = QPTrim(fptxtVendCode)
  If fpcboActive.ListIndex = 1 Then
    Vendor.ActiveFlag = 1
  Else
    Vendor.ActiveFlag = 0
  End If
30:
  Vendor.VNAME = QPTrim(fptxtVendName)
  Vendor.Addr1 = QPTrim(fptxtVendAdd1)
  Vendor.Addr2 = QPTrim(fptxtVendAdd2)
  Vendor.City = QPTrim(fptxtVendCity)
  Vendor.State = QPTrim(fptxtVendState)
  Vendor.Zip = QPTrim(fpmskVendZip)
  If Len(QPTrim(fptxtChkName)) > 0 Then
    Vendor.PaytoName = QPTrim(fptxtChkName)
  Else
    Vendor.PaytoName = QPTrim(fptxtVendName)
  End If
  Vendor.PaytoAddr = QPTrim(fptxtChkAdd1)
  Vendor.PaytoAddr2 = QPTrim(fptxtChkAdd2)
  Vendor.PayToCity = QPTrim(fptxtChkCity)
  Vendor.PaytoState = QPTrim(fptxtChkState)
  Vendor.PaytoZip = QPTrim(fpmskChkZip)
  Vendor.Memo = QPTrim(fptxtChkMemo)
  Vendor.VTerms = QPTrim(fptxtVTerms)
  Vendor.Fedid = QPTrim(fptxtVFedId)
  Vendor.StCode = QPTrim(fptxtVStateCode)
  Vendor.CoCode = QPTrim(fptxtVCountyCode)
  Vendor.Get1099 = QPTrim(Left$(fpcboVGet1099.Text, 1))
  Vendor.Contact = QPTrim(fptxtVContact)
  Vendor.Phone = QPTrim(fpmskVPhone)
  Vendor.Fax = QPTrim(fpmskVFax)
  Vendor.DBA = QPTrim(fptxtDBA)
  Vendor.ChkByte = Chr$(1)
  fpcboVendCode.col = 2
  vrec = QPTrim(fpcboVendCode.ColText)
50:
  OpenVendorFile VendorFile, NumVRecs
  If vrec = 0 Then   'VRec = 0 if called from Add New Vendor Sub
    vrec = NumVRecs + 1
  End If
60:
  Put VendorFile, vrec, Vendor
  Close VendorFile
  Call MainLog("Ol Vend Saved - " + QPTrim(fptxtVendName))
   ' IndexVendorFile frmEditVendor
saveerr:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (saveven - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
Private Sub cmdDelete_Click()
  If fpcboVendCode.ListIndex <> -1 Then
    DelRec
  End If
End Sub
Private Sub DelRec()
  Dim VendorFile As Integer, NumVRecs As Integer, vrec As Integer
  If Not Exist("APIED.dat") Then
    OpenVendorFile VendorFile, NumVRecs
    fpcboVendCode.col = 2
    vrec = QPTrim(fpcboVendCode.ColText)
    Get VendorFile, vrec, Vendor
    If Vendor.FrstTran = 0 Then
      If MsgBox("Last Chance, Delete Vendor?", vbYesNo, "Delete") = vbYes Then
        Vendor.ActiveFlag = 1
        Vendor.DelFlag = True
        'OpenVendorFile VendorFile, NumVRecs
        Put VendorFile, vrec, Vendor
        Close VendorFile
        Call MainLog("Ol Vendor Deleted - " + fpcboVendCode.ColText)
        IndexVendorFile frmEditVendor
        clearscreen
      End If
    Else
      MsgBox "This Vendor Has Transactions and May Not Be Deleted.", vbOKOnly, "Delete Canceled"
      Close VendorFile
      vaTabPro1.ActiveTab = 0
      fpcboVendCode.SetFocus
    End If
  Else
    MsgBox "UnPosted Invoices Exist And Must Be Posted Or Deleted Before Deleting A Vendor.", vbOKOnly, "Delete Vendor"
  End If
End Sub
Private Sub clearscreen()
On Local Error GoTo errstuff
  fpcboVendCode.ListIndex = -1
  fpcboActive.ListIndex = -1
  fptxtVendName = ""
  fptxtVendAdd1 = ""
  fptxtVendAdd2 = ""
  fptxtVendCity = ""
  fptxtVendState = ""
  fpmskVendZip = ""
  fptxtChkName = ""
  fptxtChkAdd1 = ""
  fptxtChkAdd2 = ""
  fptxtChkCity = ""
  fptxtChkState = ""
  fpmskChkZip = ""
  fptxtChkMemo = ""
  fptxtVTerms = 0
  fptxtVFedId = ""
  fptxtVStateCode = ""
  fptxtVCountyCode = ""
  fpcboVGet1099.ListIndex = -1
  fptxtVContact = ""
  fpmskVPhone = ""
  fpmskVFax = ""
  fptxtDBA = ""
20:
  VendCodeNameIA fpcboVendCode
21:
  VendsLstAlpha fplstVendors
22:
  vaTabPro1.ActiveTab = 0
  
  fpViewVen.Visible = False
errstuff:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (Clearscrn - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub

'******These ChangeModes are to allow editing in fp fields
Private Sub fptxtVendName_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtVendAdd1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtVendAdd2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtVendCity_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtVendState_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpmskVendZip_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtChkName_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtChkAdd1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtChkAdd2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtChkCity_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtChkState_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpmskChkZip_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtChkMemo_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtVTerms_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtVFedId_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtVStateCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtVCountyCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtVContact_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpmskVPhone_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpmskVFax_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtDBA_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
Private Sub ChkVenActivity()
  Dim APLedgerFile As Integer, NumTran As Long, RecLen As Integer
  Dim Pcnt As Integer, cnt As Integer, NextTrans As Long
  Dim VendorFile As Integer, NumVRecs As Integer, vrec As Integer
  On Local Error GoTo chkvenerr
  NumOpenItems = 0
10:
  If Not Exist("APIED.dat") And Not Exist("APPED.DAT") Then
20:
    OpenVendorFile VendorFile, NumVRecs
22:
    fpcboVendCode.col = 2
23:
    vrec = Val(QPTrim(fpcboVendCode.ColText))
24:
   If vrec > 0 Then
25:
    Get VendorFile, vrec, Vendor
26:
    If Vendor.FrstTran > 0 Then
27:
      ReDim APLedgerRec(1) As APLedger81RecType
28:
      RecLen = Len(APLedgerRec(1))
29:
      OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
30:
      NextTrans& = Vendor.FrstTran
31:
      Do Until NextTrans& = 0
32:
        Get APLedgerFile, NextTrans&, APLedgerRec(1)
33:
        If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
34:
          NumOpenItems = NumOpenItems + 1
35:
          Exit Do
36:
        End If
37:
        NextTrans& = APLedgerRec(1).NextTrans
38:
      Loop
39:
      Close
40:
      Erase APLedgerRec
41:
    End If
42:
    End If
43:
    Close
  Else
    NumOpenItems = 1
  End If
chkvenerr:
  If Err > 0 Then
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (chkven - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
End Sub

