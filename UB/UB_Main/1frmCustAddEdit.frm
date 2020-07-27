VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCustAddEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8940
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "1frmCustAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      CausesValidation=   0   'False
      Height          =   7260
      Left            =   480
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   405
      Width           =   11265
      _Version        =   196609
      _ExtentX        =   19870
      _ExtentY        =   12806
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   0
      ThreeD          =   0   'False
      TabShape        =   3
      ActiveTabBold   =   0   'False
      OffsetFromClientTop=   -1  'True
      ChamferedWidth  =   0
      ChamferedHeight =   0
      ShowEarMark     =   0   'False
      EarMarkHeight   =   0
      DataFormat      =   ""
      DataAutoRead    =   0   'False
      AutoSizeChildren=   3
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      ThreeDOuterWidth=   0
      ThreeDOuterWidthActive=   0
      ThreeDInnerWidth=   0
      ThreeDInnerWidthActive=   0
      ThreeDAppearance=   0
      DataField       =   ""
      TabCaption      =   "1frmCustAddEdit.frx":08CA
      PageEarMarkPictureNext=   "1frmCustAddEdit.frx":0B0A
      PageEarMarkPicturePrev=   "1frmCustAddEdit.frx":0B26
      EarMarkPictureNext=   "1frmCustAddEdit.frx":0B42
      EarMarkPicturePrev=   "1frmCustAddEdit.frx":0B5E
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   7065
         Left            =   45
         TabIndex        =   189
         Top             =   90
         Width           =   11160
         _Version        =   196609
         _ExtentX        =   19685
         _ExtentY        =   12462
         _StockProps     =   70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Picture         =   "1frmCustAddEdit.frx":0B7A
         Begin LpLib.fpCombo fpGroupCde 
            Height          =   330
            Left            =   8640
            TabIndex        =   4
            Top             =   1875
            Width           =   2325
            _Version        =   196608
            _ExtentX        =   4101
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
            Columns         =   3
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
            ThreeDOutsideStyle=   2
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   1
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   1
            EditMarginRight =   0
            EditMarginBottom=   3
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   -1  'True
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":0B96
         End
         Begin LpLib.fpCombo fpBillTo 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   8550
            TabIndex        =   20
            Top             =   4515
            Width           =   1560
            _Version        =   196608
            _ExtentX        =   2752
            _ExtentY        =   556
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
            Columns         =   1
            Sorted          =   1
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
            DataSync        =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2388
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   0   'False
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":0F49
         End
         Begin LpLib.fpCombo fpStatus 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   6120
            TabIndex        =   2
            Top             =   1440
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   2
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2388
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   0   'False
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":12A4
         End
         Begin EditLib.fpMask fpZip 
            Height          =   300
            Left            =   6816
            TabIndex        =   12
            Top             =   4032
            Width           =   1548
            _Version        =   196608
            _ExtentX        =   2730
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpDateTime fpOpenDate 
            Height          =   300
            Left            =   8640
            TabIndex        =   3
            Top             =   1464
            Width           =   1692
            _Version        =   196608
            _ExtentX        =   2984
            _ExtentY        =   529
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
            ButtonStyle     =   3
            ButtonWidth     =   23
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   0
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm-dd-yyyy"
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin fpBtnAtlLibCtl.fpBtn btnPages 
            Height          =   324
            Left            =   9672
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   216
            Width           =   1356
            _Version        =   131072
            _ExtentX        =   2392
            _ExtentY        =   572
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":176E
         End
         Begin EditLib.fpText fpBook 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   0
            Top             =   1464
            Width           =   372
            _Version        =   196608
            _ExtentX        =   656
            _ExtentY        =   529
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
            AlignTextV      =   1
            AllowNull       =   -1  'True
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
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
         Begin EditLib.fpText fpSeqNumb 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   3240
            TabIndex        =   1
            Top             =   1464
            Width           =   876
            _Version        =   196608
            _ExtentX        =   1545
            _ExtentY        =   529
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
            AlignTextV      =   1
            AllowNull       =   -1  'True
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "0123456789"
            MaxLength       =   6
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
         Begin fpBtnAtlLibCtl.fpBtn btnPageInfo 
            Height          =   495
            Left            =   270
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   210
            Width           =   3840
            _Version        =   131072
            _ExtentX        =   6773
            _ExtentY        =   873
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":194C
         End
         Begin EditLib.fpText fpSearch 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   5
            Top             =   2112
            Width           =   1548
            _Version        =   196608
            _ExtentX        =   2730
            _ExtentY        =   529
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "~ "
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
         Begin EditLib.fpText fpCustName 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   6
            Top             =   2568
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpAddr1 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   7
            Top             =   2928
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpAddr2 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   8
            Top             =   3288
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpServAddr 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   9
            Top             =   3648
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpCity 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   10
            Top             =   4032
            Width           =   2100
            _Version        =   196608
            _ExtentX        =   3704
            _ExtentY        =   529
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
            MarginTop       =   0
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
            MaxLength       =   18
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
         Begin EditLib.fpText fpState 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   5736
            TabIndex        =   11
            Top             =   4032
            Width           =   420
            _Version        =   196608
            _ExtentX        =   741
            _ExtentY        =   529
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "~-0123456789"
            MaxLength       =   2
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
         Begin EditLib.fpMask fpHPhone 
            Height          =   300
            Left            =   2664
            TabIndex        =   14
            Top             =   4560
            Width           =   1620
            _Version        =   196608
            _ExtentX        =   2857
            _ExtentY        =   529
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
            MarginTop       =   0
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
            Mask            =   "(###) ###-####"
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
         Begin EditLib.fpMask fpWPhone 
            Height          =   300
            Left            =   2664
            TabIndex        =   15
            Top             =   4872
            Width           =   1620
            _Version        =   196608
            _ExtentX        =   2857
            _ExtentY        =   529
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
            MarginTop       =   0
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
            Mask            =   "(###) ###-####"
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
         Begin EditLib.fpMask fpSoSec 
            Height          =   300
            Left            =   2664
            TabIndex        =   16
            Top             =   5208
            Width           =   1620
            _Version        =   196608
            _ExtentX        =   2857
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpBillCopy 
            Height          =   300
            Left            =   8544
            TabIndex        =   21
            Top             =   4848
            Width           =   348
            _Version        =   196608
            _ExtentX        =   614
            _ExtentY        =   529
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "1"
            CharValidationText=   "1234567890"
            MaxLength       =   2
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
         Begin EditLib.fpText fpPostRte 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   8544
            TabIndex        =   22
            Top             =   5184
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpBillCycl 
            Height          =   300
            Left            =   8544
            TabIndex        =   23
            Top             =   5544
            Width           =   348
            _Version        =   196608
            _ExtentX        =   614
            _ExtentY        =   529
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
            MaxLength       =   2
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
         Begin EditLib.fpText fpZone 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   8544
            TabIndex        =   24
            Top             =   5880
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
            _ExtentY        =   529
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
            MarginTop       =   0
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
            MaxLength       =   3
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
         Begin EditLib.fpText fpDrvLic 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   17
            Top             =   5544
            Width           =   2100
            _Version        =   196608
            _ExtentX        =   3704
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpCustType 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   18
            Top             =   5880
            Width           =   636
            _Version        =   196608
            _ExtentX        =   1122
            _ExtentY        =   529
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
            MarginTop       =   0
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
            MaxLength       =   3
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
         Begin EditLib.fpText fpAddr911 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   19
            Top             =   6240
            Width           =   2100
            _Version        =   196608
            _ExtentX        =   3704
            _ExtentY        =   529
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
            MarginTop       =   0
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
            MaxLength       =   14
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
         Begin EditLib.fpText fpSeq 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   8544
            TabIndex        =   25
            Top             =   6240
            Width           =   900
            _Version        =   196608
            _ExtentX        =   1587
            _ExtentY        =   529
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
            AlignTextV      =   1
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "0123456789"
            MaxLength       =   8
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
         Begin EditLib.fpText fpDPCode 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   9456
            TabIndex        =   13
            Top             =   4032
            Width           =   636
            _Version        =   196608
            _ExtentX        =   1122
            _ExtentY        =   529
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
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
         Begin EditLib.fpText fpPrnBillYN 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   8640
            TabIndex        =   324
            Top             =   2310
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
            _ExtentY        =   529
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
            AlignTextV      =   1
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "yYnN"
            MaxLength       =   1
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
         Begin VB.Label Label104 
            Alignment       =   1  'Right Justify
            Caption         =   "Print Bill:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7218
            TabIndex        =   323
            Top             =   2355
            Width           =   1290
         End
         Begin VB.Label Label103 
            Alignment       =   1  'Right Justify
            Caption         =   "Group Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6864
            TabIndex        =   322
            Top             =   1920
            Width           =   1644
         End
         Begin VB.Label Label102 
            Caption         =   "DP Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   8568
            TabIndex        =   321
            Top             =   4056
            Width           =   852
         End
         Begin VB.Label LabelAcctNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   276
            Left            =   8568
            TabIndex        =   26
            Top             =   864
            Width           =   1140
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Acct No:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7320
            TabIndex        =   159
            Top             =   864
            Width           =   1140
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Read Sequence No:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   154
            Top             =   6264
            Width           =   2220
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Zone Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   153
            Top             =   5904
            Width           =   2220
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Billing Cycle:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   152
            Top             =   5568
            Width           =   2220
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Postal Sort:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   151
            Top             =   5208
            Width           =   2220
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Bill Copies:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   150
            Top             =   4872
            Width           =   2220
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Bill To:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6216
            TabIndex        =   149
            Top             =   4560
            Width           =   2220
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "911 Address:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   148
            Top             =   6264
            Width           =   2220
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   147
            Top             =   5904
            Width           =   2220
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Drivers Licenses No:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   146
            Top             =   5568
            Width           =   2220
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Social Security Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   96
            TabIndex        =   145
            Top             =   5208
            Width           =   2460
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Work Phone:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   144
            Top             =   4872
            Width           =   2220
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Home Phone:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   143
            Top             =   4560
            Width           =   2220
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Zip:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6192
            TabIndex        =   142
            Top             =   4032
            Width           =   564
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "State:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4968
            TabIndex        =   141
            Top             =   4032
            Width           =   684
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "City:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   140
            Top             =   4032
            Width           =   2220
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Service Address:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   139
            Top             =   3648
            Width           =   2220
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   138
            Top             =   3288
            Width           =   2220
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Address Line 1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   137
            Top             =   2928
            Width           =   2220
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   136
            Top             =   2568
            Width           =   2220
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Search Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   336
            TabIndex        =   135
            Top             =   2112
            Width           =   2220
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Open Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7368
            TabIndex        =   134
            Top             =   1512
            Width           =   1140
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   5064
            TabIndex        =   133
            Top             =   1488
            Width           =   996
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3072
            TabIndex        =   132
            Top             =   1416
            Width           =   132
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Location Number:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   384
            TabIndex        =   131
            Top             =   1488
            Width           =   2220
         End
      End
      Begin ImpproLib.vaImprint vaImprint4 
         Height          =   7005
         Left            =   -26190
         TabIndex        =   191
         Top             =   -22020
         Width           =   11190
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
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
         Caption         =   ""
         Picture         =   "1frmCustAddEdit.frx":1B37
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   0
            Left            =   2895
            TabIndex        =   63
            Top             =   3840
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":1B53
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   0
            Left            =   2205
            TabIndex        =   62
            Top             =   3840
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":1FD7
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   1
            Left            =   2205
            TabIndex        =   72
            Top             =   4200
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":25CC
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   2
            Left            =   2205
            TabIndex        =   82
            Top             =   4590
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":2BC1
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   3
            Left            =   2205
            TabIndex        =   92
            Top             =   4965
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":31B6
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   4
            Left            =   2205
            TabIndex        =   102
            Top             =   5355
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":37AB
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   5
            Left            =   2205
            TabIndex        =   112
            Top             =   5730
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":3DA0
         End
         Begin LpLib.fpCombo fpLocMType 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   6
            Left            =   2205
            TabIndex        =   122
            Top             =   6120
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":4395
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   1
            Left            =   2895
            TabIndex        =   73
            Top             =   4200
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":498A
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   2
            Left            =   2895
            TabIndex        =   83
            Top             =   4590
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":4E0E
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   3
            Left            =   2895
            TabIndex        =   93
            Top             =   4965
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":5292
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   4
            Left            =   2895
            TabIndex        =   103
            Top             =   5355
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":5716
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   5
            Left            =   2895
            TabIndex        =   113
            Top             =   5730
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":5B9A
         End
         Begin LpLib.fpCombo fpLocUnit 
            CausesValidation=   0   'False
            Height          =   315
            Index           =   6
            Left            =   2895
            TabIndex        =   123
            Top             =   6120
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":601E
         End
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   0
            Left            =   5424
            TabIndex        =   66
            Top             =   3840
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   0
            Left            =   4116
            TabIndex        =   65
            Top             =   3840
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   4
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   0
            Left            =   1668
            TabIndex        =   61
            Top             =   3840
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin fpBtnAtlLibCtl.fpBtn fpBtn4 
            Height          =   324
            Left            =   9672
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   216
            Width           =   1356
            _Version        =   131072
            _ExtentX        =   2392
            _ExtentY        =   572
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":64A2
         End
         Begin fpBtnAtlLibCtl.fpBtn fpBtn6 
            Height          =   495
            Left            =   270
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   210
            Width           =   3840
            _Version        =   131072
            _ExtentX        =   6773
            _ExtentY        =   873
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":6680
         End
         Begin EditLib.fpCurrency fpMonOwed 
            Height          =   312
            Index           =   0
            Left            =   984
            TabIndex        =   49
            Top             =   1800
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMonOwed 
            Height          =   312
            Index           =   1
            Left            =   984
            TabIndex        =   53
            Top             =   2160
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMonPaid 
            Height          =   312
            Index           =   0
            Left            =   2136
            TabIndex        =   50
            Top             =   1800
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMonPaid 
            Height          =   312
            Index           =   1
            Left            =   2136
            TabIndex        =   54
            Top             =   2184
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMonAmt 
            Height          =   312
            Index           =   0
            Left            =   3240
            TabIndex        =   51
            Top             =   1800
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMonAmt 
            Height          =   312
            Index           =   1
            Left            =   3240
            TabIndex        =   55
            Top             =   2184
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpLongInteger fpMonRev 
            Height          =   315
            Index           =   0
            Left            =   4410
            TabIndex        =   52
            Top             =   1800
            Width           =   465
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   550
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMonRev 
            Height          =   312
            Index           =   1
            Left            =   4392
            TabIndex        =   56
            Top             =   2184
            Width           =   468
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   550
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpCurrency fpMemFee 
            Height          =   312
            Index           =   0
            Left            =   6768
            TabIndex        =   57
            Top             =   1800
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpMemFee 
            Height          =   312
            Index           =   1
            Left            =   8568
            TabIndex        =   58
            Top             =   1800
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   550
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "$0.00"
            CurrencyDecimalPlaces=   -1
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   ""
            DecimalPoint    =   ""
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   0
            Left            =   216
            TabIndex        =   60
            Top             =   3840
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   1
            Left            =   216
            TabIndex        =   70
            Top             =   4200
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   2
            Left            =   216
            TabIndex        =   80
            Top             =   4584
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   3
            Left            =   216
            TabIndex        =   90
            Top             =   4968
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   4
            Left            =   216
            TabIndex        =   100
            Top             =   5352
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   5
            Left            =   216
            TabIndex        =   110
            Top             =   5736
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpMtrSerial 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   6
            Left            =   216
            TabIndex        =   120
            Top             =   6120
            Width           =   1452
            _Version        =   196608
            _ExtentX        =   2561
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   1
            Left            =   1668
            TabIndex        =   71
            Top             =   4200
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   2
            Left            =   1668
            TabIndex        =   81
            Top             =   4584
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   3
            Left            =   1668
            TabIndex        =   91
            Top             =   4968
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   4
            Left            =   1668
            TabIndex        =   101
            Top             =   5352
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   5
            Left            =   1668
            TabIndex        =   111
            Top             =   5736
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrMulti 
            Height          =   312
            Index           =   6
            Left            =   1668
            TabIndex        =   121
            Top             =   6120
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   0
            Left            =   3576
            TabIndex        =   64
            Top             =   3840
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   1
            Left            =   3576
            TabIndex        =   74
            Top             =   4200
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   2
            Left            =   3576
            TabIndex        =   84
            Top             =   4584
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   3
            Left            =   3576
            TabIndex        =   94
            Top             =   4968
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   4
            Left            =   3576
            TabIndex        =   104
            Top             =   5352
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   5
            Left            =   3576
            TabIndex        =   114
            Top             =   5736
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpMtrUser 
            Height          =   312
            Index           =   6
            Left            =   3576
            TabIndex        =   124
            Top             =   6120
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   1
            Left            =   4116
            TabIndex        =   75
            Top             =   4200
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   2
            Left            =   4116
            TabIndex        =   85
            Top             =   4584
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   3
            Left            =   4116
            TabIndex        =   95
            Top             =   4968
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   4
            Left            =   4116
            TabIndex        =   105
            Top             =   5352
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   5
            Left            =   4116
            TabIndex        =   115
            Top             =   5736
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMtrIns 
            Height          =   312
            Index           =   6
            Left            =   4116
            TabIndex        =   125
            Top             =   6120
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   1
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   0
            Left            =   7944
            TabIndex        =   68
            Top             =   3840
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   4
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   1
            Left            =   7944
            TabIndex        =   78
            Top             =   4200
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   2
            Left            =   7944
            TabIndex        =   88
            Top             =   4584
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   3
            Left            =   7944
            TabIndex        =   98
            Top             =   4968
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   4
            Left            =   7944
            TabIndex        =   108
            Top             =   5352
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   5
            Left            =   7944
            TabIndex        =   118
            Top             =   5736
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpLocMLRDate 
            Height          =   312
            Index           =   6
            Left            =   7944
            TabIndex        =   128
            Top             =   6120
            Width           =   1308
            _Version        =   196608
            _ExtentX        =   2307
            _ExtentY        =   550
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   2
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   3
            NullColor       =   -2147483643
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   3
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
            Appearance      =   0
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
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   1
            Left            =   5424
            TabIndex        =   76
            Top             =   4200
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   2
            Left            =   5424
            TabIndex        =   86
            Top             =   4584
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   3
            Left            =   5424
            TabIndex        =   96
            Top             =   4968
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   4
            Left            =   5424
            TabIndex        =   106
            Top             =   5352
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   5
            Left            =   5424
            TabIndex        =   116
            Top             =   5736
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrCur 
            Height          =   312
            Index           =   6
            Left            =   5424
            TabIndex        =   126
            Top             =   6120
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   0
            Left            =   6684
            TabIndex        =   67
            Top             =   3840
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   1
            Left            =   6684
            TabIndex        =   77
            Top             =   4200
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   2
            Left            =   6684
            TabIndex        =   87
            Top             =   4584
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   3
            Left            =   6684
            TabIndex        =   97
            Top             =   4968
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   4
            Left            =   6684
            TabIndex        =   107
            Top             =   5352
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   5
            Left            =   6684
            TabIndex        =   117
            Top             =   5736
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpLocMtrPre 
            Height          =   312
            Index           =   6
            Left            =   6684
            TabIndex        =   127
            Top             =   6120
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   550
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
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "1234567890"
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   0
            Left            =   9264
            TabIndex        =   69
            Top             =   3840
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   1
            Left            =   9264
            TabIndex        =   79
            Top             =   4200
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   2
            Left            =   9264
            TabIndex        =   89
            Top             =   4584
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   3
            Left            =   9264
            TabIndex        =   99
            Top             =   4968
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   4
            Left            =   9264
            TabIndex        =   109
            Top             =   5352
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   5
            Left            =   9264
            TabIndex        =   119
            Top             =   5736
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin EditLib.fpText fpMtrIDNO 
            CausesValidation=   0   'False
            Height          =   312
            Index           =   6
            Left            =   9264
            TabIndex        =   129
            Top             =   6120
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   550
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
            CaretOverWrite  =   0
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   11
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
         Begin fpBtnAtlLibCtl.fpBtn fpcmdMtrCoordinates 
            Height          =   384
            Left            =   4500
            TabIndex        =   59
            Top             =   6528
            Width           =   2172
            _Version        =   131072
            _ExtentX        =   3831
            _ExtentY        =   677
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":686B
         End
         Begin VB.Label Label101 
            Alignment       =   2  'Center
            Caption         =   "ID Info"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   9504
            TabIndex        =   221
            Top             =   3528
            Width           =   1212
         End
         Begin VB.Label Label93 
            Alignment       =   2  'Center
            Caption         =   "Reading"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   6768
            TabIndex        =   206
            Top             =   3528
            Width           =   1092
         End
         Begin VB.Label Label98 
            Alignment       =   2  'Center
            Caption         =   "Last Read"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   7992
            TabIndex        =   205
            Top             =   3264
            Width           =   1212
         End
         Begin VB.Label Label97 
            Alignment       =   2  'Center
            Caption         =   "Previous"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   6768
            TabIndex        =   204
            Top             =   3264
            Width           =   1092
         End
         Begin VB.Label Label96 
            Alignment       =   2  'Center
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   5544
            TabIndex        =   203
            Top             =   3264
            Width           =   1020
         End
         Begin VB.Label Label95 
            Alignment       =   2  'Center
            Caption         =   "Installed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   4176
            TabIndex        =   202
            Top             =   3264
            Width           =   1164
         End
         Begin VB.Label Label94 
            Alignment       =   2  'Center
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   8208
            TabIndex        =   201
            Top             =   3528
            Width           =   780
         End
         Begin VB.Label Label92 
            Alignment       =   2  'Center
            Caption         =   "Reading"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   5520
            TabIndex        =   200
            Top             =   3528
            Width           =   1092
         End
         Begin VB.Label Label91 
            Alignment       =   2  'Center
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   4368
            TabIndex        =   199
            Top             =   3528
            Width           =   852
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            Caption         =   "Usr"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   3600
            TabIndex        =   198
            Top             =   3552
            Width           =   492
         End
         Begin VB.Label Label89 
            Alignment       =   2  'Center
            Caption         =   "Unt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   2880
            TabIndex        =   197
            Top             =   3552
            Width           =   516
         End
         Begin VB.Label Label88 
            Alignment       =   2  'Center
            Caption         =   "Typ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   2256
            TabIndex        =   196
            Top             =   3528
            Width           =   468
         End
         Begin VB.Label Label87 
            Alignment       =   2  'Center
            Caption         =   "Mul"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1680
            TabIndex        =   195
            Top             =   3528
            Width           =   564
         End
         Begin VB.Label Label86 
            Alignment       =   2  'Center
            Caption         =   "Serial #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   264
            Left            =   504
            TabIndex        =   194
            Top             =   3504
            Width           =   1092
         End
         Begin VB.Label Label85 
            Alignment       =   1  'Right Justify
            Caption         =   "Meter Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   192
            TabIndex        =   218
            Top             =   2760
            Width           =   2292
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            Caption         =   "Non-Refundable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   8136
            TabIndex        =   217
            Top             =   1488
            Width           =   1716
         End
         Begin VB.Label Label83 
            Alignment       =   1  'Right Justify
            Caption         =   "Refundable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6528
            TabIndex        =   215
            Top             =   1488
            Width           =   1260
         End
         Begin VB.Label Label82 
            Alignment       =   2  'Center
            Caption         =   "Membership Fees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   6744
            TabIndex        =   216
            Top             =   840
            Width           =   2556
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            Caption         =   "Rev."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   4272
            TabIndex        =   214
            Top             =   1488
            Width           =   588
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3168
            TabIndex        =   213
            Top             =   1488
            Width           =   972
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Caption         =   "Paid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   2064
            TabIndex        =   212
            Top             =   1488
            Width           =   732
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            Caption         =   "Owed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   936
            TabIndex        =   211
            Top             =   1488
            Width           =   804
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1920
            TabIndex        =   210
            Top             =   1272
            Width           =   1044
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   840
            TabIndex        =   209
            Top             =   1272
            Width           =   1020
         End
         Begin VB.Label Label60 
            Alignment       =   2  'Center
            Caption         =   "Monthly Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   1656
            TabIndex        =   208
            Top             =   840
            Width           =   2700
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   7005
         Left            =   -26190
         TabIndex        =   190
         Top             =   -22020
         Width           =   11190
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
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
         Caption         =   ""
         Picture         =   "1frmCustAddEdit.frx":6A4F
         Begin EditLib.fpLongInteger fpProRatePCT 
            Height          =   300
            Left            =   8568
            TabIndex        =   43
            Top             =   1824
            Width           =   588
            _Version        =   196608
            _ExtentX        =   1037
            _ExtentY        =   529
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
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "100"
            MaxValue        =   "100"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin fpBtnAtlLibCtl.fpBtn fpBtn2 
            Height          =   324
            Left            =   9672
            TabIndex        =   48
            Top             =   216
            Width           =   1356
            _Version        =   131072
            _ExtentX        =   2392
            _ExtentY        =   572
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":6A6B
         End
         Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
            Height          =   495
            Left            =   270
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   210
            Width           =   3840
            _Version        =   131072
            _ExtentX        =   6773
            _ExtentY        =   873
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":6C49
         End
         Begin EditLib.fpBoolean fpLateFee 
            Height          =   300
            Left            =   2664
            TabIndex        =   28
            Top             =   1800
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   1
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean fpCutOffYN 
            Height          =   300
            Left            =   2664
            TabIndex        =   29
            Top             =   2136
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   1
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean fpTaxExpt 
            Height          =   300
            Left            =   2664
            TabIndex        =   30
            Top             =   2472
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean fpSrCit 
            Height          =   300
            Left            =   2664
            TabIndex        =   31
            Top             =   2808
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpBoolean fpUseDraft 
            Height          =   300
            Left            =   2664
            TabIndex        =   32
            Top             =   3480
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin EditLib.fpText fpAcctType 
            Height          =   300
            Left            =   2664
            TabIndex        =   33
            Top             =   3816
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
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
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   "CcSs"
            MaxLength       =   1
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
         Begin EditLib.fpText fpBankName 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   34
            Top             =   4152
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
            MaxLength       =   34
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
         Begin EditLib.fpText fpBankLoc 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   35
            Top             =   4488
            Width           =   3924
            _Version        =   196608
            _ExtentX        =   6921
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpTransit 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   36
            Top             =   4824
            Width           =   1164
            _Version        =   196608
            _ExtentX        =   2053
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpBankAcct 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   37
            Top             =   5160
            Width           =   2364
            _Version        =   196608
            _ExtentX        =   4170
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpBillCmnt 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   38
            Top             =   5880
            Width           =   2892
            _Version        =   196608
            _ExtentX        =   5101
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpPayCmnt 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   2664
            TabIndex        =   39
            Top             =   6216
            Width           =   2892
            _Version        =   196608
            _ExtentX        =   5101
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpPumpCode 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   6120
            TabIndex        =   40
            Top             =   1488
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpUserCode1 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   6120
            TabIndex        =   41
            Top             =   1824
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpUserCode2 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   6120
            TabIndex        =   42
            Top             =   2160
            Width           =   372
            _Version        =   196608
            _ExtentX        =   656
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpHHMsg1 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   7608
            TabIndex        =   44
            Top             =   3504
            Width           =   2364
            _Version        =   196608
            _ExtentX        =   4170
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpHHMsg2 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   7608
            TabIndex        =   45
            Top             =   3816
            Width           =   2364
            _Version        =   196608
            _ExtentX        =   4170
            _ExtentY        =   529
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
            MarginTop       =   0
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
         Begin EditLib.fpText fpHHMsg3 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   7605
            TabIndex        =   46
            Top             =   4140
            Width           =   2370
            _Version        =   196608
            _ExtentX        =   4170
            _ExtentY        =   529
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
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   0
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
         Begin EditLib.fpBoolean fpCashOnly 
            Height          =   300
            Left            =   2664
            TabIndex        =   27
            Top             =   1464
            Width           =   324
            _Version        =   196608
            _ExtentX        =   572
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            AutoToggle      =   -1  'True
            BooleanStyle    =   1
            ToggleFalse     =   "Nn"
            TextFalse       =   "N"
            BooleanPicture  =   0
            AlignPictureH   =   3
            AlignPictureV   =   1
            GroupId         =   0
            GroupTag        =   0
            GroupSelect     =   0
            MarginLeft      =   3
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            MultiLine       =   0   'False
            AlignTextH      =   1
            AlignTextV      =   0
            ToggleTrue      =   "Yy"
            TextTrue        =   "Y"
            Value           =   0
            BooleanMode     =   0
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            BorderGrayAreaColor=   -2147483637
            ToggleGrayed    =   ""
            TextGrayed      =   ""
            AllowMnemonic   =   0   'False
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDOnFocusInvert=   0   'False
            Caption         =   "N"
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            BooleanDataType =   0
            OLEDropMode     =   0
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7176
            TabIndex        =   187
            Top             =   4128
            Width           =   324
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7176
            TabIndex        =   186
            Top             =   3816
            Width           =   324
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            Caption         =   "1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7176
            TabIndex        =   185
            Top             =   3504
            Width           =   324
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            Caption         =   "Handheld Message/Comment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7224
            TabIndex        =   184
            Top             =   3144
            Width           =   3036
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "Prorate:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7536
            TabIndex        =   183
            Top             =   1848
            Width           =   948
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "User Code 2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4464
            TabIndex        =   182
            Top             =   2184
            Width           =   1596
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "User Code 1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4464
            TabIndex        =   181
            Top             =   1848
            Width           =   1596
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "Source Pump:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4464
            TabIndex        =   180
            Top             =   1512
            Width           =   1596
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Payment Comment:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   179
            Top             =   6240
            Width           =   2220
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "Billing Comment:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   178
            Top             =   5904
            Width           =   2220
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Account No:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   177
            Top             =   5208
            Width           =   2220
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Transit No:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   176
            Top             =   4848
            Width           =   2220
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Location:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   175
            Top             =   4512
            Width           =   2220
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   174
            Top             =   4176
            Width           =   2220
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Account Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   173
            Top             =   3840
            Width           =   2220
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Draft Account (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   172
            Top             =   3528
            Width           =   2220
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Sr. Citizen (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   171
            Top             =   2856
            Width           =   2220
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax Exempt (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   170
            Top             =   2520
            Width           =   2220
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow Cut Off (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   169
            Top             =   2184
            Width           =   2220
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow Late Fee (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   168
            Top             =   1848
            Width           =   2220
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Cash Only (Y/N):"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   167
            Top             =   1512
            Width           =   2220
         End
      End
      Begin ImpproLib.vaImprint vaImprint3 
         Height          =   7005
         Left            =   -26190
         TabIndex        =   222
         Top             =   -22020
         Width           =   11190
         _Version        =   196609
         _ExtentX        =   19748
         _ExtentY        =   12361
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
         Caption         =   ""
         Picture         =   "1frmCustAddEdit.frx":6E34
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   0
            Left            =   3312
            TabIndex        =   223
            Top             =   1560
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   0   'False
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":6E50
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   1
            Left            =   3315
            TabIndex        =   225
            Top             =   1875
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":71D7
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   2
            Left            =   3315
            TabIndex        =   227
            Top             =   2190
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":755E
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   3
            Left            =   3312
            TabIndex        =   229
            Top             =   2496
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":78E5
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   4
            Left            =   3312
            TabIndex        =   231
            Top             =   2808
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":7C6C
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   5
            Left            =   3312
            TabIndex        =   233
            Top             =   3120
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":7FF3
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   6
            Left            =   3315
            TabIndex        =   235
            Top             =   3435
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   2
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":837A
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   7
            Left            =   3315
            TabIndex        =   237
            Top             =   3750
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":8701
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   8
            Left            =   8250
            TabIndex        =   239
            Top             =   1530
            Width           =   930
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":8A88
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   9
            Left            =   8256
            TabIndex        =   241
            Top             =   1872
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1640
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":8E0F
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   10
            Left            =   8256
            TabIndex        =   243
            Top             =   2184
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1640
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":9196
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   11
            Left            =   8256
            TabIndex        =   247
            Top             =   2496
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":951D
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   12
            Left            =   8256
            TabIndex        =   249
            Top             =   2808
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":98A4
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   13
            Left            =   8256
            TabIndex        =   251
            Top             =   3120
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1630
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":9C2B
         End
         Begin LpLib.fpCombo fpServCode 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   14
            Left            =   8256
            TabIndex        =   254
            Top             =   3432
            Width           =   924
            _Version        =   196608
            _ExtentX        =   1640
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   1
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   0
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   0
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   1
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   3000
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   4
            MaxEditLen      =   0
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
            AutoSearchFillDelay=   500
            EditMarginLeft  =   0
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   -1  'True
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":9FB2
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   0
            Left            =   4410
            TabIndex        =   224
            Top             =   1560
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":A339
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   1
            Left            =   4416
            TabIndex        =   226
            Top             =   1872
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":A92E
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   2
            Left            =   4416
            TabIndex        =   228
            Top             =   2184
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":AF23
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   3
            Left            =   4416
            TabIndex        =   230
            Top             =   2496
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":B518
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   4
            Left            =   4416
            TabIndex        =   232
            Top             =   2808
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":BB0D
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   5
            Left            =   4416
            TabIndex        =   234
            Top             =   3120
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":C102
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   6
            Left            =   4416
            TabIndex        =   236
            Top             =   3432
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":C6F7
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   7
            Left            =   4416
            TabIndex        =   238
            Top             =   3744
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":CCEC
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   8
            Left            =   9360
            TabIndex        =   240
            Top             =   1560
            Width           =   690
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":D2E1
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   9
            Left            =   9360
            TabIndex        =   242
            Top             =   1872
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":D8D6
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   10
            Left            =   9360
            TabIndex        =   245
            Top             =   2184
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":DECB
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   11
            Left            =   9360
            TabIndex        =   248
            Top             =   2496
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":E4C0
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   12
            Left            =   9360
            TabIndex        =   250
            Top             =   2808
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":EAB5
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   13
            Left            =   9360
            TabIndex        =   253
            Top             =   3120
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1206
            _ExtentY        =   466
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":F0AA
         End
         Begin LpLib.fpCombo fpServMType 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   14
            Left            =   9360
            TabIndex        =   255
            Top             =   3432
            Width           =   684
            _Version        =   196608
            _ExtentX        =   1217
            _ExtentY        =   476
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
            Columns         =   2
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   1
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":F69F
         End
         Begin LpLib.fpCombo fpFlatFreq 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   0
            Left            =   5760
            TabIndex        =   258
            Top             =   5328
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   466
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
            Columns         =   1
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   0
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":FC94
         End
         Begin LpLib.fpCombo fpFlatFreq 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   1
            Left            =   5760
            TabIndex        =   264
            Top             =   5664
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   476
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
            Columns         =   1
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   0
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":FFEF
         End
         Begin LpLib.fpCombo fpFlatFreq 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   2
            Left            =   5760
            TabIndex        =   274
            Top             =   6000
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   466
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
            Columns         =   1
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   0
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":1034A
         End
         Begin LpLib.fpCombo fpFlatFreq 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   3
            Left            =   5760
            TabIndex        =   284
            Top             =   6336
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
            _ExtentY        =   466
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
            Columns         =   1
            Sorted          =   0
            SelDrawFocusRect=   -1  'True
            ColumnSeparatorChar=   9
            ColumnSearch    =   0
            ColumnWidthScale=   3
            RowHeight       =   -1
            WrapList        =   0   'False
            WrapWidth       =   0
            AutoSearch      =   1
            SearchMethod    =   0
            VirtualMode     =   0   'False
            VRowCount       =   0
            DataSync        =   0
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
            ScrollBarV      =   3
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
            DataAutoSizeCols=   0
            SearchIgnoreCase=   -1  'True
            ScrollBarH      =   3
            DataFieldList   =   ""
            ColumnEdit      =   0
            ColumnBound     =   -1
            Style           =   2
            MaxDrop         =   8
            ListWidth       =   2580
            EditHeight      =   -1
            GrayAreaColor   =   -2147483633
            ListLeftOffset  =   0
            ComboGap        =   0
            MaxEditLen      =   0
            VirtualPageSize =   0
            VirtualPagesAhead=   0
            ExtendCol       =   2
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
            ExtendRow       =   2
            ListPosition    =   0
            ButtonThreeDAppearance=   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            Redraw          =   -1  'True
            AutoSearchFill  =   -1  'True
            AutoSearchFillDelay=   500
            EditMarginLeft  =   2
            EditMarginTop   =   0
            EditMarginRight =   0
            EditMarginBottom=   0
            ResizeRowToFont =   0   'False
            TextTipMultiLine=   0
            AutoMenu        =   0   'False
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "1frmCustAddEdit.frx":106A5
         End
         Begin EditLib.fpLongInteger fpFlatMin 
            Height          =   264
            Index           =   0
            Left            =   9264
            TabIndex        =   260
            Top             =   5328
            Width           =   420
            _Version        =   196608
            _ExtentX        =   741
            _ExtentY        =   466
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatRevSrc 
            Height          =   264
            Index           =   0
            Left            =   8112
            TabIndex        =   259
            Top             =   5328
            Width           =   468
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   462
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpCurrency fpFlatAmt 
            Height          =   264
            Index           =   0
            Left            =   4104
            TabIndex        =   257
            Top             =   5328
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   466
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
            ButtonWrap      =   0   'False
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
            MarginTop       =   0
            MarginRight     =   0
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
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin fpBtnAtlLibCtl.fpBtn fpBtn3 
            Height          =   324
            Left            =   9672
            TabIndex        =   244
            TabStop         =   0   'False
            Top             =   216
            Width           =   1356
            _Version        =   131072
            _ExtentX        =   2392
            _ExtentY        =   572
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":10A00
         End
         Begin fpBtnAtlLibCtl.fpBtn fpBtn5 
            Height          =   495
            Left            =   270
            TabIndex        =   246
            TabStop         =   0   'False
            Top             =   210
            Width           =   3840
            _Version        =   131072
            _ExtentX        =   6773
            _ExtentY        =   873
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":10BDE
         End
         Begin EditLib.fpText fpFlatDesc 
            CausesValidation=   0   'False
            Height          =   270
            Index           =   0
            Left            =   1260
            TabIndex        =   256
            Top             =   5310
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3789
            _ExtentY        =   466
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpFlatDesc 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   1
            Left            =   1248
            TabIndex        =   261
            Top             =   5664
            Width           =   2148
            _Version        =   196608
            _ExtentX        =   3789
            _ExtentY        =   466
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpFlatDesc 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   2
            Left            =   1248
            TabIndex        =   270
            Top             =   6000
            Width           =   2148
            _Version        =   196608
            _ExtentX        =   3789
            _ExtentY        =   466
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin EditLib.fpText fpFlatDesc 
            CausesValidation=   0   'False
            Height          =   264
            Index           =   3
            Left            =   1248
            TabIndex        =   280
            Top             =   6336
            Width           =   2148
            _Version        =   196608
            _ExtentX        =   3789
            _ExtentY        =   466
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
            MarginTop       =   0
            MarginRight     =   0
            MarginBottom    =   0
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   18
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
         Begin fpBtnAtlLibCtl.fpBtn fpBtn7 
            Height          =   420
            Left            =   504
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   4464
            Width           =   3228
            _Version        =   131072
            _ExtentX        =   5694
            _ExtentY        =   741
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   0   'False
            GrayAreaColor   =   12632256
            BorderShowDefault=   -1  'True
            ButtonType      =   0
            NoPointerFocus  =   -1  'True
            Value           =   0   'False
            GroupID         =   0
            GroupSelect     =   0
            DrawFocusRect   =   2
            DrawFocusRectCell=   -1
            GrayAreaPictureStyle=   0
            Static          =   -1  'True
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
            ButtonDesigner  =   "1frmCustAddEdit.frx":10DC9
         End
         Begin EditLib.fpCurrency fpFlatAmt 
            Height          =   264
            Index           =   1
            Left            =   4104
            TabIndex        =   262
            Top             =   5664
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   466
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
            ButtonWrap      =   0   'False
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
            MarginTop       =   0
            MarginRight     =   0
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
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpFlatAmt 
            Height          =   270
            Index           =   2
            Left            =   4110
            TabIndex        =   272
            Top             =   6000
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   466
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
            ButtonWrap      =   0   'False
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
            MarginTop       =   0
            MarginRight     =   0
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
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpCurrency fpFlatAmt 
            Height          =   264
            Index           =   3
            Left            =   4104
            TabIndex        =   282
            Top             =   6336
            Width           =   972
            _Version        =   196608
            _ExtentX        =   1714
            _ExtentY        =   466
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
            ButtonWrap      =   0   'False
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
            MarginTop       =   0
            MarginRight     =   0
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
            MaxValue        =   "9999.99"
            MinValue        =   "0"
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
         Begin EditLib.fpLongInteger fpFlatRevSrc 
            Height          =   264
            Index           =   1
            Left            =   8112
            TabIndex        =   266
            Top             =   5664
            Width           =   468
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   462
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatRevSrc 
            Height          =   264
            Index           =   2
            Left            =   8112
            TabIndex        =   276
            Top             =   6000
            Width           =   468
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   462
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatRevSrc 
            Height          =   264
            Index           =   3
            Left            =   8112
            TabIndex        =   286
            Top             =   6336
            Width           =   468
            _Version        =   196608
            _ExtentX        =   825
            _ExtentY        =   462
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "15"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatMin 
            Height          =   264
            Index           =   1
            Left            =   9264
            TabIndex        =   268
            Top             =   5664
            Width           =   420
            _Version        =   196608
            _ExtentX        =   741
            _ExtentY        =   466
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatMin 
            Height          =   264
            Index           =   2
            Left            =   9264
            TabIndex        =   278
            Top             =   6000
            Width           =   420
            _Version        =   196608
            _ExtentX        =   741
            _ExtentY        =   466
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger fpFlatMin 
            Height          =   264
            Index           =   3
            Left            =   9264
            TabIndex        =   288
            Top             =   6336
            Width           =   420
            _Version        =   196608
            _ExtentX        =   741
            _ExtentY        =   466
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
            MarginLeft      =   0
            MarginTop       =   0
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1392
            TabIndex        =   320
            Top             =   1152
            Width           =   1284
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3336
            TabIndex        =   319
            Top             =   1152
            Width           =   588
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   7
            Left            =   1176
            TabIndex        =   318
            Top             =   3744
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1170
            TabIndex        =   317
            Top             =   3450
            UseMnemonic     =   0   'False
            Width           =   1950
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   5
            Left            =   1176
            TabIndex        =   316
            Top             =   3120
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   4
            Left            =   1176
            TabIndex        =   315
            Top             =   2808
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   3
            Left            =   1176
            TabIndex        =   314
            Top             =   2496
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   2
            Left            =   1176
            TabIndex        =   313
            Top             =   2184
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   1
            Left            =   1176
            TabIndex        =   312
            Top             =   1872
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   0
            Left            =   1176
            TabIndex        =   311
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            Caption         =   "1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   310
            Top             =   1560
            Width           =   348
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            Caption         =   "2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   309
            Top             =   1872
            Width           =   348
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            Caption         =   "3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   308
            Top             =   2184
            Width           =   348
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            Caption         =   "4)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   307
            Top             =   2496
            Width           =   348
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            Caption         =   "5)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   306
            Top             =   2808
            Width           =   348
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Caption         =   "6)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   305
            Top             =   3120
            Width           =   348
         End
         Begin VB.Label Label73 
            Alignment       =   1  'Right Justify
            Caption         =   "7)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   304
            Top             =   3432
            Width           =   348
         End
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            Caption         =   "8)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   303
            Top             =   3744
            Width           =   348
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            Caption         =   "15)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   302
            Top             =   3432
            Width           =   348
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Caption         =   "14)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   301
            Top             =   3120
            Width           =   348
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "13)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   300
            Top             =   2808
            Width           =   348
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            Caption         =   "12)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   299
            Top             =   2496
            Width           =   348
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            Caption         =   "11)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   298
            Top             =   2184
            Width           =   348
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            Caption         =   "10)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   297
            Top             =   1872
            Width           =   348
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            Caption         =   "9)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5592
            TabIndex        =   296
            Top             =   1560
            Width           =   348
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   14
            Left            =   6096
            TabIndex        =   295
            Top             =   3432
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   13
            Left            =   6096
            TabIndex        =   294
            Top             =   3120
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   12
            Left            =   6096
            TabIndex        =   293
            Top             =   2808
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   11
            Left            =   6096
            TabIndex        =   292
            Top             =   2496
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   10
            Left            =   6096
            TabIndex        =   291
            Top             =   2184
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   9
            Left            =   6096
            TabIndex        =   290
            Top             =   1872
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label PG3RevLBL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   8
            Left            =   6096
            TabIndex        =   289
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   1956
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   8136
            TabIndex        =   287
            Top             =   1152
            Width           =   708
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6264
            TabIndex        =   285
            Top             =   1152
            Width           =   1356
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1368
            TabIndex        =   283
            Top             =   4968
            Width           =   1356
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            Caption         =   "4)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   281
            Top             =   6360
            Width           =   348
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            Caption         =   "3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   279
            Top             =   6024
            Width           =   348
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            Caption         =   "2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   277
            Top             =   5688
            Width           =   348
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            Caption         =   "1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   744
            TabIndex        =   275
            Top             =   5352
            Width           =   348
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3936
            TabIndex        =   273
            Top             =   4968
            Width           =   1020
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "Frequency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   5832
            TabIndex        =   271
            Top             =   4968
            Width           =   1236
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            Caption         =   "Minimum"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   8808
            TabIndex        =   269
            Top             =   4968
            Width           =   1140
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            Caption         =   "Revenue"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   7680
            TabIndex        =   267
            Top             =   4968
            Width           =   1068
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
            Caption         =   "Mtr Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4056
            TabIndex        =   265
            Top             =   1152
            Width           =   1116
         End
         Begin VB.Label Label100 
            Alignment       =   1  'Right Justify
            Caption         =   "Mtr Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   8952
            TabIndex        =   263
            Top             =   1176
            Width           =   1116
         End
      End
   End
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   768
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   144
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
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
      Text            =   ""
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   96
      Top             =   144
   End
   Begin fpBtnAtlLibCtl.fpBtn btnPgUp 
      Height          =   384
      Left            =   11616
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   7728
      Width           =   444
      _Version        =   131072
      _ExtentX        =   783
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":10FB2
   End
   Begin fpBtnAtlLibCtl.fpBtn btnPgDn 
      Height          =   384
      Left            =   11616
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   8088
      Width           =   444
      _Version        =   131072
      _ExtentX        =   783
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":1336C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   9048
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":15726
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   10320
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":15901
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOwner 
      Height          =   384
      Left            =   6504
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":15ADC
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdConHist 
      Height          =   384
      Left            =   3960
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":15CB7
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdTranHist 
      Height          =   384
      Left            =   2688
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":15E93
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdWorkHist 
      Height          =   390
      Left            =   1410
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":1606F
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrintInfo 
      Height          =   384
      Left            =   144
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":1624B
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   193
      Top             =   8610
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   582
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
            TextSave        =   "10:40 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "3/12/2020"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdMsg 
      Height          =   384
      Left            =   5232
      TabIndex        =   219
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":16427
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdWOE 
      Height          =   390
      Left            =   7770
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":16600
   End
End
Attribute VB_Name = "frmCustAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim RecNo As Long, CntL As Long
Dim TransRec As Long, MsgRec As Long
Dim UBSetUpRec(1) As UBSetupRecType
Dim UBOwnerRec As UBOwnerRecType
Dim UBSetupLen As Integer, cnt As Integer
Dim OldBook As String, NBook As String
Dim FinalFlag As Boolean, UpDateOwner As Boolean
Dim BeenDone As Boolean
Dim BtnFnt As Double
Dim fromform As Form, toform As Form, codeopt As Integer
Dim dontdoit As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub Form_Load()
  BlockInput True
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  CustAddEdLoadRateCodes
  setbillto
  setfreq
  
  FillGroupCMBO fpGroupCde
  DoEvents
  dontdoit = False
  BlockInput False
  Hook fpServCode(0).hWnd
  
  If Not GetAllowEBills Then
  'If InStr(TOWNNAME$, "DEEP RUN") Or InStr(TOWNNAME$, "CLINTWOOD") < 1 Then
    fpPrnBillYN.Visible = False
    Label104.Visible = False
  End If
  
  DoEvents
End Sub
Private Sub setbillto()
  fpBillTo.AddItem "Customer"
  fpBillTo.AddItem "Owner"
  fpBillTo.ListIndex = 0
End Sub

Private Sub setfreq()
  Dim cnt As Integer
  For cnt = 0 To 3
    fpFlatFreq(cnt).AddItem "Recurring"
    fpFlatFreq(cnt).AddItem "NonRecurring"
    fpFlatFreq(cnt).ListIndex = -1
  Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via CustAddEdit by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    'DoEvents
    Temp_Class.ResizeControls Me
   ' DoEvents
   ' Me.Visible = True
   ' Me.AutoRedraw = False
   ' DoEvents
  End If
  DoEvents
End Sub

Private Sub Form_Activate()
  BlockInput True
  If Val(frmCustAddEdit.fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    LoadCustInfo2Form
    DoEvents
  ElseIf Val(frmCustAddEdit.fpCustRecNo) <= 0 And Not BeenDone Then
    DoEvents
    fpCmdTranHist.Enabled = False
    fpcmdPrintinfo.Enabled = False
    fpCmdWorkHist.Enabled = False
    fpCmdConHist.Enabled = False
    fpCmdMsg.Enabled = False
    fpcmdMtrCoordinates.Enabled = False
    NewCustDefaults
    setup4new
  End If
  BlockInput False
End Sub
'
'Mouse/Keyboard/Button events
'
Private Sub NewCustDefaults()
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  fpCity = QPTrim(UBSetUpRec(1).DEFCITY)
  fpState = QPTrim(UBSetUpRec(1).DEFSTATE)
  fpZip = QPTrim(UBSetUpRec(1).ZIPCODE)
  fpStatus.ListIndex = 0
  fpOpenDate = Format(Now, "mm/dd/yyyy")
  fpBillTo.ListIndex = 0
  fpBillCopy = 1
  For cnt = 0 To 6
    fpMtrMulti(cnt) = 1
    fpMtrUser(cnt) = 1
  Next
  fpGroupCde.ListIndex = 0

 UBLog PWUser + " New Cust Entry"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyPageDown
    '  If Not ListIsDown Then
        KeyCode = 0
        If vaTabPro1.ActiveTab < 3 Then
          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab + 1
        Else
          vaTabPro1.ActiveTab = 0
        End If
      'End If
    Case vbKeyPageUp
'      If Not ListIsDown Then
        KeyCode = 0
        If vaTabPro1.ActiveTab > 0 Then
          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab - 1
        Else
          vaTabPro1.ActiveTab = 3
        End If
'      Else
'        KeyCode = 0
'      End If
    Case vbKeyF2
      KeyCode = 0
      Call fpcmdPrintinfo_Click
    Case vbKeyF3
      KeyCode = 0
      Call fpCmdWorkHist_Click
    Case vbKeyF4
      KeyCode = 0
      Call fpCmdTranHist_Click
      'trans history
    Case vbKeyF6
      KeyCode = 0
      Call fpCmdConHist_Click
    Case vbKeyF7
      KeyCode = 0
      Call fpCmdMsg_Click
    Case vbKeyF8
      KeyCode = 0
      Call fpCmdOwner_Click
    Case vbKeyF9
      KeyCode = 0
      Call fpCmdWOE_Click
    Case vbKeyF10
      KeyCode = 0
      DoEvents
      Call fpCmdSave_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnHook
End Sub

Private Sub fpcmdPrintinfo_Click()
  If RecNo& > 0 Then
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
    'do the graphics
      PrintCustInfo RecNo&, 1
    ElseIf rptopt = 2 Then
    'do the text
      PrintCustInfo RecNo&, 2
    End If
   ActivateControls Me
  Else
    ActivateControls Me
  End If

End Sub

Private Sub btnPgUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    Sendkeys "{PgUp}", True
  End If
End Sub

Private Sub btnPgDn_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    Sendkeys "{PgDn}", True
  End If
End Sub

'---------------------------------
'&&&&&& Page 1 Keydowns
'---------------------------------

Private Sub fpBook_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpSeqNumb.SetFocus
    Case vbKeyUp
      KeyCode = 0
      Sendkeys "{PgUp}"  'This does not work for first tab
  End Select
End Sub

Private Sub fpcmdMtrCoordinates_Click()
  Dim tmpCustRec As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  CustRecLen = Len(tmpCustRec)
  
  UBHandle = FreeFile
  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
  UBLog PWUser + (" Opened Custfile, fpcmdMtrCoordinates, for - " + Str(RecNo&) + " with " + Str(CustRecLen) + " len ")
  Get #UBHandle, RecNo&, tmpCustRec
  Close UBHandle
  UBLog PWUser + (" Closed Custfile, fpcmdMtrCoordinates ")
  For cnt = 0 To 6
    frmCustMtrCoordinates.fpLatitude(cnt) = tmpCustRec.LocMeters(cnt + 1).MtrLat
    frmCustMtrCoordinates.fpLongitude(cnt) = tmpCustRec.LocMeters(cnt + 1).MtrLng
  Next
  DoEvents
  frmCustMtrCoordinates.fpCustRecNo = RecNo
  frmCustMtrCoordinates.Show 1
End Sub

Private Sub fpSeqNumb_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpStatus.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBook.SetFocus
  End Select
End Sub


Private Sub fpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpStatus.ListIndex = -1
    fpStatus.Action = ActionClearSearchBuffer
  End If
  If fpStatus.ListDown <> True Then
    If KeyCode = vbKeySpace Then
      KeyCode = 0
      fpStatus.ListDown = True
    End If
    If KeyCode = vbKeyUp Then
      KeyCode = 0
      If Me.fpSeqNumb.Enabled = True Then
        Me.fpSeqNumb.SetFocus
      End If
    Else
      If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        KeyCode = 0
        Me.fpOpenDate.SetFocus
      End If
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpStatus.ListDown = False
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpOpenDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpGroupCde.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpStatus.SetFocus
    End Select
End Sub
Private Sub fpGroupCde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpGroupCde.ListIndex = -1
    fpGroupCde.Action = ActionClearSearchBuffer
  End If
  If fpGroupCde.ListDown <> True Then
    If KeyCode = vbKeySpace Then
      KeyCode = 0
      fpGroupCde.ListDown = True
    End If
    If KeyCode = vbKeyUp Then
      KeyCode = 0
      If Me.fpOpenDate.Enabled = True Then
        Me.fpOpenDate.SetFocus
      End If
    Else
      If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
        KeyCode = 0
        Me.fpSearch.SetFocus
      End If
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpGroupCde.ListDown = False
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpCustName.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpGroupCde.SetFocus
    End Select
End Sub
Private Sub fpCustName_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpAddr1.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpSearch.SetFocus
    End Select
End Sub
Private Sub fpAddr1_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpAddr2.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpCustName.SetFocus
    End Select
End Sub
Private Sub fpAddr2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpServAddr.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpAddr1.SetFocus
    End Select
End Sub

Private Sub fpServAddr_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpCity.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpAddr2.SetFocus
    End Select
End Sub
Private Sub fpCity_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpState.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpServAddr.SetFocus
    End Select
End Sub
Private Sub fpState_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpZip.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpCity.SetFocus
    End Select
End Sub

Private Sub fpstatus_SelChange(ItemIndex As Long)
  Dim CStatus As String
  CStatus = QPTrim$(Me.fpStatus.Text)
  If CStatus <> "F" Then
    fpBook.Enabled = True
    fpSeqNumb.Enabled = True
  End If
End Sub

Private Sub fpZip_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpDPCode.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpState.SetFocus
    End Select
End Sub
Private Sub fpDPCode_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpHPhone.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpZip.SetFocus
    End Select
End Sub

Private Sub fpHPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpWPhone.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpDPCode.SetFocus
    End Select
End Sub
Private Sub fpWPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpSoSec.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpHPhone.SetFocus
    End Select
End Sub
Private Sub fpSoSec_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpDrvLic.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpWPhone.SetFocus
    End Select
End Sub
Private Sub fpDrvLic_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpCustType.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpSoSec.SetFocus
    End Select
End Sub
Private Sub fpCustType_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpAddr911.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpDrvLic.SetFocus
    End Select
End Sub
Private Sub fpAddr911_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBillTo.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpCustType.SetFocus
    End Select
End Sub
Private Sub fpBillTo_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not fpBillTo.ListDown Then
    Select Case KeyCode
    Case vbKeySpace
      fpBillTo.ListDown = True
    Case vbKeyUp
      KeyCode = 0
      fpAddr911.SetFocus
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBillCopy.SetFocus
    End Select
  Else
    If KeyCode = vbKeySpace Then
      fpBillTo.ListDown = False
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpBillCopy_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpPostRte.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBillTo.SetFocus
    End Select
End Sub
Private Sub fpPostRte_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBillCycl.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBillCopy.SetFocus
    End Select
End Sub
Private Sub fpBillCycl_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpZone.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpPostRte.SetFocus
    End Select
End Sub
Private Sub fpZone_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpSeq.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBillCycl.SetFocus
    End Select
End Sub

Private Sub fpSeq_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      Sendkeys "{PgDn}"
    Case vbKeyUp
      KeyCode = 0
      fpZone.SetFocus
  End Select
End Sub
'----------------------------
'&&&&&&&&&   Page 2 keydowns
'------------------------------
Private Sub fpCashOnly_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpLateFee.SetFocus
    Case vbKeyUp, vbKeyBack
      KeyCode = 0
      Sendkeys "{PgUp}"
  End Select
End Sub
Private Sub fpLateFee_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpCutOffYN.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpCashOnly.SetFocus
    End Select
End Sub
Private Sub fpCutOffYN_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpTaxExpt.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpLateFee.SetFocus
    End Select
End Sub
Private Sub fpTaxExpt_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpSrCit.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpCutOffYN.SetFocus
    End Select
End Sub
Private Sub fpSrCit_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpUseDraft.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpTaxExpt.SetFocus
    End Select
End Sub
Private Sub fpUseDraft_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpAcctType.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpSrCit.SetFocus
    End Select
End Sub
Private Sub fpAcctType_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBankName.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpUseDraft.SetFocus
    End Select
End Sub
Private Sub fpBankName_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBankLoc.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpAcctType.SetFocus
    End Select
End Sub
Private Sub fpBankLoc_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpTransit.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBankName.SetFocus
    End Select
End Sub
Private Sub fpTransit_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBankAcct.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBankLoc.SetFocus
    End Select
End Sub
Private Sub fpBankAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpBillCmnt.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpTransit.SetFocus
    End Select
End Sub
Private Sub fpBillCmnt_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpPayCmnt.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBankAcct.SetFocus
    End Select
End Sub
Private Sub fpPayCmnt_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpPumpCode.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpBillCmnt.SetFocus
    End Select
End Sub
Private Sub fpPumpCode_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpUserCode1.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpPayCmnt.SetFocus
    End Select
End Sub
Private Sub fpUserCode1_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpUserCode2.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpPumpCode.SetFocus
    End Select
End Sub
Private Sub fpUserCode2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpProRatePCT.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpUserCode1.SetFocus
    End Select
End Sub
Private Sub fpProRatePCT_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpHHMsg1.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpUserCode2.SetFocus
    End Select
End Sub
Private Sub fpHHMsg1_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpHHMsg2.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpProRatePCT.SetFocus
    End Select
End Sub
Private Sub fpHHMsg2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn
      KeyCode = 0
      fpHHMsg3.SetFocus
    Case vbKeyUp
      KeyCode = 0
      fpHHMsg1.SetFocus
    End Select
End Sub
Private Sub fpHHMsg3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    KeyCode = 0
    Sendkeys "{PgDn}"
  ElseIf KeyCode = vbKeyUp Then
      KeyCode = 0
      fpHHMsg2.SetFocus
  End If
End Sub
'---------------------------------
'&&&&&&&&&&&   Page 3 keydowns
'--------------------------------
Private Sub fpServCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpServCode(Index).ListIndex = -1
    fpServCode(Index).Action = ActionClearSearchBuffer
    KeyCode = 0
  End If
  If fpServCode(Index).ListDown <> True Then
    If KeyCode = vbKeySpace Then
      fpServCode(Index).ListDown = True
      KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      If Index <> 14 Then
        If fpServMType(Index).Enabled = True Then
          fpServMType(Index).SetFocus
        Else
          If fpServCode(Index + 1).Enabled = True Then
            fpServCode(Index + 1).SetFocus
          Else
            fpFlatDesc(0).SetFocus
          End If
        End If
        KeyCode = 0
      Else
        fpFlatDesc(0).SetFocus
        KeyCode = 0
      End If
    Else
      If KeyCode = vbKeyUp Then
        If Index <> 0 Then
          If fpServMType(Index - 1).Enabled = True Then
            fpServMType(Index - 1).SetFocus
          Else
            If fpServCode(Index - 1).Enabled = True Then
              fpServCode(Index - 1).SetFocus
            Else
              Sendkeys "{PgUp}"
            End If
          End If
          KeyCode = 0
        Else
          KeyCode = 0
          Sendkeys "{PgUp}"
        End If
       End If
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpServCode(Index).ListDown = False
      KeyCode = 0
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If

  End If
  DoEvents
  
End Sub
Private Sub fpServMType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpServMType(Index).ListIndex = -1
    fpServMType(Index).Action = ActionClearSearchBuffer
    KeyCode = 0
  End If
  If fpServMType(Index).ListDown <> True Then
    If KeyCode = vbKeySpace Then
      fpServMType(Index).ListDown = True
      KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      If Index <> 14 Then
        If fpServCode(Index + 1).Enabled = True Then
          fpServCode(Index + 1).SetFocus
        Else
          fpFlatDesc(0).SetFocus
        End If
      Else
        fpFlatDesc(0).SetFocus
      End If
      KeyCode = 0
    End If
    If KeyCode = vbKeyUp Then
      If Index <> 0 Then
        If fpServCode(Index).Enabled = True Then
          fpServCode(Index).SetFocus
        ElseIf fpServMType(Index - 1).Enabled = True Then
          fpServMType(Index - 1).SetFocus
        ElseIf fpServCode(Index - 1).Enabled = True Then
          fpServCode(Index - 1).SetFocus
        End If
        KeyCode = 0
      Else
        If fpServCode(Index).Enabled = True Then
          fpServCode(Index).SetFocus
        Else
          Sendkeys "{PgUp}"
        End If
        KeyCode = 0
      End If
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpServMType(Index).ListDown = False
      KeyCode = 0
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
  DoEvents
End Sub
Private Sub fpFlatDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyUp
    If Index <> 0 Then
      fpFlatMin(Index - 1).SetFocus
      KeyCode = 0
    Else
      For cnt = 14 To 0
        If fpServCode(cnt).Enabled = True Then
          fpServCode(cnt).SetFocus
          KeyCode = 0
          Exit For
        End If
      Next
    End If
  Case vbKeyReturn, vbKeyDown, vbKeyTab
    fpFlatAmt(Index).SetFocus
    KeyCode = 0
  Case Else
  End Select
End Sub
Private Sub fpFlatAmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn, vbKeyTab
    fpFlatFreq(Index).SetFocus
    KeyCode = 0
  Case vbKeyUp
    fpFlatDesc(Index).SetFocus
    KeyCode = 0
  End Select
End Sub
Private Sub fpFlatFreq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    fpFlatFreq(Index).ListIndex = -1
    fpFlatFreq(Index).Action = ActionClearSearchBuffer
    KeyCode = 0
  End If
  If fpFlatFreq(Index).ListDown <> True Then
    If KeyCode = vbKeySpace Then
      fpFlatFreq(Index).ListDown = True
      KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
      fpFlatRevSrc(Index).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      fpFlatAmt(Index).SetFocus
      KeyCode = 0
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpFlatFreq(Index).ListDown = False
      KeyCode = 0
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpFlatRevSrc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyDown, vbKeyReturn, vbKeyTab
    fpFlatMin(Index).SetFocus
    KeyCode = 0
  Case vbKeyUp
    fpFlatFreq(Index).SetFocus
    KeyCode = 0
  End Select
End Sub
Private Sub fpFlatMin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If Index <> 3 Then
      fpFlatDesc(Index + 1).SetFocus
    Else
      Sendkeys "{PgDn}"
    End If
    KeyCode = 0
 ElseIf KeyCode = vbKeyUp Then
    fpFlatRevSrc(Index).SetFocus
    KeyCode = 0
 End If
End Sub
'----------------------------------------
'&&&&&&&&&&&     Page 4 keydowns
'----------------------------------------
Private Sub fpMonOwed_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    KeyCode = 0
    fpMonPaid(Index).SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If Index = 0 Then
      KeyCode = 0
      Sendkeys "{PgUp}"
    Else
      fpMonRev(Index - 1).SetFocus
    End If
  End If
End Sub
Private Sub fpMonPaid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    KeyCode = 0
    fpMonAmt(Index).SetFocus
  ElseIf KeyCode = vbKeyUp Then
    KeyCode = 0
    fpMonOwed(Index).SetFocus
  End If
End Sub
Private Sub fpMonAmt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    KeyCode = 0
    fpMonRev(Index).SetFocus
  ElseIf KeyCode = vbKeyUp Then
    KeyCode = 0
    fpMonPaid(Index).SetFocus
  End If
End Sub
Private Sub fpMonRev_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    If Index = 1 Then
      fpMemFee(0).SetFocus
    Else
      fpMonOwed(Index + 1).SetFocus
    End If
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    KeyCode = 0
    fpMonAmt(Index).SetFocus
  End If
End Sub
Private Sub fpMemFee_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    If Index = 0 Then
      fpMemFee(1).SetFocus
      KeyCode = 0
    Else
      fpMtrSerial(0).SetFocus
      KeyCode = 0
    End If
  ElseIf KeyCode = vbKeyUp Then
    If Index = 0 Then
      fpMonRev(1).SetFocus
      KeyCode = 0
    Else
      fpMemFee(0).SetFocus
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpMtrSerial_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpMtrMulti(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    If Index = 0 Then
      fpMemFee(1).SetFocus
    Else
      fpMtrIDNO(Index - 1).SetFocus
    End If
    KeyCode = 0
  End If
  
 ' KeyCode = 0
End Sub
Private Sub fpMtrMulti_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpLocMType(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpMtrSerial(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpLocMType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If fpLocMType(Index).ListDown = False Then
    If KeyCode = vbKeySpace Then
      fpLocMType(Index).ListDown = True
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpLocUnit(Index).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      fpMtrMulti(Index).SetFocus
      KeyCode = 0
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpLocMType(Index).ListDown = False
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub
Private Sub fpLocUnit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If fpLocUnit(Index).ListDown = False Then
    If KeyCode = vbKeySpace Then
      fpLocUnit(Index).ListDown = True
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpMtrUser(Index).SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      fpLocMType(Index).SetFocus
      KeyCode = 0
    End If
  Else
    If KeyCode = vbKeySpace Then
      fpLocUnit(Index).ListDown = False
      KeyCode = 0
    End If
    If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpMtrUser_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpLocMtrIns(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpLocUnit(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpLocMtrIns_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpLocMtrCur(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpMtrUser(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpLocMtrCur_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpLocMtrPre(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpLocMtrIns(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpLocMtrPre_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpLocMLRDate(Index).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyUp Then
    fpLocMtrCur(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpLocMLRDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    fpMtrIDNO(Index).SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpLocMtrPre(Index).SetFocus
    KeyCode = 0
  End If
End Sub
Private Sub fpMtrIDNO_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
    If Index = 6 Then
      Sendkeys "{PgDn}"
      KeyCode = 0
    Else
      fpMtrSerial(Index + 1).SetFocus
      KeyCode = 0
    End If
  ElseIf KeyCode = vbKeyUp Then
    fpLocMLRDate(Index).SetFocus
    KeyCode = 0
  End If
End Sub
    
Private Sub fpCmdConHist_Click()
  If RecNo > 0 Then
    If Exist(UBPath$ + "UBTRANS.DAT") Then
      frmRptConsumpHist.ShowCustConsHist (RecNo&)
    Else
      MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
    End If
  End If
End Sub

Private Sub fpCmdWOE_Click()
  If RecNo& <= 0 Then
  'need to give option to save
    If MsgBox("You must save new customer info before entering workorders.", vbOKCancel, "Save info?") = vbCancel Then
      Exit Sub
    Else
      If ChkCustInfoOK% Then
        If dontdoit = False Then
          Call SaveCustInfo2Disk
          dontdoit = False
        End If
        WorkOrders
      End If
    End If
  Else
    Select Case CheckSaveCustFile%
    Case True:  '-1 save chenges
    If ChkCustInfoOK% Then
      If dontdoit = False Then
        Call SaveCustInfo2Disk
        dontdoit = False
      End If
      WorkOrders
    End If
    Case False:  '0= exit
      WorkOrders
    Case Else     '1 is review
      'stay right where you are
    End Select
  End If
End Sub
Private Sub fpMtrSerial_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub
Private Sub fpMtrIDNO_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCmdWorkHist_Click()
  If RecNo > 0 Then
    frmRptWrkOrdHist.ShowWrkOrdHistory (RecNo&)
  End If
End Sub
Private Sub fpStatus_LostFocus()
  Dim CStatus As String
  CStatus = QPTrim$(Me.fpStatus.Text)
  If Len(CStatus) = 0 Then
    MsgBox "   Customer Status CAN NOT BE BLANK!   ", vbOKOnly, "ERROR"
    vaTabPro1.ActiveTab = 0
    DoEvents
    Me.fpStatus.SetFocus
  End If
End Sub



Private Sub vaTabPro1_TabPageShown(ActiveTab As Integer, ActivePage As Integer)
  On Local Error GoTo skimover
  Select Case ActiveTab
  Case 0
    If fpBook.Enabled = True Then
      Me.fpBook.SetFocus
    Else
      Me.fpStatus.SetFocus
    End If
  Case 1
    Me.fpCashOnly.SetFocus
  Case 2
    Me.fpServCode(0).SetFocus
  Case 3
    Me.fpMonOwed(0).SetFocus
  End Select
skimover:
End Sub
'END Mouse/Keyboard/Button Sections

Private Sub fpCmdMsg_Click()
  If RecNo > 0 Then
    frmCustMsgEdit.CustRec = RecNo
    frmCustMsgEdit.Show vbModal
    DoEvents
    If CustHasMsg(RecNo) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      fpCmdMsg.ForeColor = &H80000012
      'fpCmdMsg.FontSize = BtnFnt
    End If
  End If
End Sub

Private Sub fpCmdOwner_Click()
  frmCustOwnerEdit.RecNo = RecNo
  frmCustOwnerEdit.Show vbModal
  DoEvents
  UpDateOwner = frmCustOwnerEdit.ActionFlag
  If UpDateOwner And RecNo > 0 Then  'an existing cust account
    Call UBSaveOwnerInfo(RecNo)      'update owner info now. (user may not update cust)
    UpDateOwner = False
  End If                        'hay, Just forget about it.
  DoEvents
  'Call UNLoadOwnerForm
  'Unload frmCustOwnerEdit
End Sub
'Private Sub UNLoadOwnerForm()
'  Unload frmCustOwnerEdit
'  DoEvents
'End Sub

Private Sub UBSaveOwnerInfo(OwnerRecNo As Long)
  Dim UBFile As Integer, OwnerRecLen As Integer
  OwnerRecLen = Len(UBOwnerRec)
  UBOwnerRec.OwnFName = frmCustOwnerEdit.fpFirstName  'new owner info until user
  UBOwnerRec.OwnLName = frmCustOwnerEdit.fpLastName   'saves new cust account.
  UBOwnerRec.ADDR1 = frmCustOwnerEdit.fpAddr1
  UBOwnerRec.ADDR2 = frmCustOwnerEdit.fpAddr2
  UBOwnerRec.CITY = frmCustOwnerEdit.fpCity
  UBOwnerRec.STATE = frmCustOwnerEdit.fpState
  UBOwnerRec.ZIPCODE = frmCustOwnerEdit.fpZip
  UBOwnerRec.HPHONE = frmCustOwnerEdit.fpHPhone
  UBOwnerRec.WPHONE = frmCustOwnerEdit.fpWPhone
  UBOwnerRec.ChkByte = Chr$(1)
  UBFile = FreeFile
  Open UBOwnerFile For Random Shared As UBFile Len = OwnerRecLen
  Put UBFile, OwnerRecNo, UBOwnerRec
  Close UBFile
End Sub

Private Sub fpCmdSave_Click()  'f10
 If ChkCustInfoOK% Then
    If dontdoit = False Then
      Call SaveCustInfo2Disk
      If codeopt = 0 Then
        NewCustDefaults
        setup4new
      Else
        Call ExitCustAddEdit
      End If
    End If
  End If
End Sub

Private Sub fpCmdExit_Click()
  Select Case CheckSaveCustFile%
  Case True:  '-1 save changes
  If ChkCustInfoOK% Then
    If dontdoit = False Then
      Call SaveCustInfo2Disk
    End If
    Call ExitCustAddEdit
  End If
  Case False:  '0= exit
'    ExitingForm = True
    Call ExitCustAddEdit
  Case Else     '1 is review
    'continue editing
  End Select
End Sub
'F7
'Display customer transaction history
Private Sub fpCmdTranHist_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If TransRec > 0 Then
    'DeActivateControls Me
    DisplayCustTransList RecNo
    'ActivateControls Me
    Select Case vaTabPro1.ActiveTab
    Case 0
      If Me.fpBook.Enabled = True Then
        Me.fpBook.SetFocus
      Else
        Me.fpStatus.SetFocus
      End If
    Case 1
      Me.fpCashOnly.SetFocus
    Case 2
      Me.fpServCode(0).SetFocus
    Case 3
      Me.fpMonOwed(0).SetFocus
    End Select
  Else
  MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(2).FontSize
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = ""
'    MsgText(3) = "There are NO transactions to display."
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
  End If
End Sub

Private Sub fpSeqNumb_LostFocus()
  Call ChkFormatBookSeqN
End Sub

Private Sub ExitCustAddEdit()
On Local Error Resume Next
  DoEvents
  RecNo = 0
  BeenDone = False
  TransRec = 0
  fpCustRecNo = 0
  NBook$ = ""
  MsgRec = 0
  OldBook = ""
  FinalFlag = False
  UpDateOwner = False
'  Load frmUBCustMenu
'  DoEvents
'  frmUBCustMenu.Show
  DoEvents
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If
  If codeopt = 0 Then
    frmUBCustMenu.Show
  End If
  UBLog PWUser + " Exit CustAddEdit"
  Unload frmCustAddEdit
  Unload frmCustOwnerEdit
  'Call UNLoadOwnerForm
'  DoEvents
End Sub

'Private Function ListIsDown%()
'  ListIsDown = False
'  If fpstatus.ListDown Then
'    ListIsDown = True
'    GoTo ExitListIsDown
'  End If
'  For cnt = 0 To 14
'    If fpServMType(cnt).ListDown Or fpServCode(cnt).ListDown Then
'      ListIsDown = True
'      Exit For
'    End If
'  Next
'ExitListIsDown:
'End Function

Private Sub SaveCustInfo2Disk()
 ' Dim ClearRFlag As Boolean
  DeActivateControls frmCustAddEdit
  
  ReDim tmpCustRec(1 To 2) As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  Dim ReindexFlag As Boolean
  Dim NextCRec As Long
  Dim PBYN As String * 1
  BlockInput True
 ' dontdoit = True
  CustRecLen = Len(tmpCustRec(1))
  If RecNo& > 0 Then
    UBHandle = FreeFile
    Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
    UBLog PWUser + (" Opened Custfile, SaveCust/AddEdit, for - " + Str(RecNo&) + " with " + Str(CustRecLen) + " len ")
    Get #UBHandle, RecNo&, tmpCustRec(1)
    Close UBHandle
    UBLog PWUser + (" Closed Custfile, SaveCust/AddEdit, Got Rec# " + Str(RecNo&))
    LSet tmpCustRec(2) = tmpCustRec(1) 'copy for reindex comparison check below
  End If
'  If tmpCustRec(2).Status <> "A" Then
'    ClearRFlag = True
'  Else
'    ClearRFlag = True
'  End If
  tmpCustRec(1).Book = QPTrim$(fpBook.Text)
  tmpCustRec(1).SEQNUMB = QPTrim$(fpSeqNumb.Text)
  tmpCustRec(1).Status = QPTrim$(fpStatus.Text)
  tmpCustRec(1).OPENDATE = Date2Num(fpOpenDate.Text)
  tmpCustRec(1).SEARCH = QPTrim$(fpSearch.Text)
  tmpCustRec(1).CustName = QPTrim$(fpCustName.Text)
  tmpCustRec(1).ADDR1 = QPTrim$(fpAddr1.Text)
  tmpCustRec(1).ADDR2 = QPTrim$(fpAddr2.Text)
    
  tmpCustRec(1).ServAddr = QPTrim$(fpServAddr.Text)
  tmpCustRec(1).CITY = QPTrim$(fpCity.Text)
  tmpCustRec(1).STATE = QPTrim$(fpState.Text)
'check
  tmpCustRec(1).ZIPCODE = QPTrim$(fpZip.Text)
  tmpCustRec(1).DPCode = QPTrim$(fpDPCode.Text)
  tmpCustRec(1).HPHONE = QPTrim$(fpHPhone.Text)
  tmpCustRec(1).WPHONE = QPTrim$(fpWPhone.Text)
  tmpCustRec(1).SOSEC = QPTrim$(fpSoSec.Text)
  tmpCustRec(1).DRVLIC = QPTrim$(fpDrvLic.Text)
  tmpCustRec(1).CUSTTYPE = QPTrim$(fpCustType.Text)
  tmpCustRec(1).Addr911 = QPTrim$(fpAddr911.Text)
  If fpBillTo.ListIndex = 1 Then
    tmpCustRec(1).BillTo = "O"
  Else
    tmpCustRec(1).BillTo = "C"
  End If
  tmpCustRec(1).BILLCOPY = Val(fpBillCopy.Text)
  tmpCustRec(1).POSTRTE = QPTrim$(fpPostRte.Text)
  If Len(QPTrim$(fpBillCycl.Text)) = 0 Then
    tmpCustRec(1).BILLCYCL = -32767
  Else
    tmpCustRec(1).BILLCYCL = Val(fpBillCycl.Text)
  End If
  tmpCustRec(1).ZONE = QPTrim$(fpZone.Text)
  If Len(QPTrim$(fpSeq.Text)) = 0 Then
    tmpCustRec(1).Seq = -32767
  Else
    tmpCustRec(1).Seq = Val(fpSeq.Text)
  End If
  fpGroupCde.col = 0
  tmpCustRec(1).GroupCodeRec = Val(fpGroupCde.ColText)
  tmpCustRec(1).CASHONLY = fpCashOnly.Text
  tmpCustRec(1).LATEFEE = fpLateFee.Text
  tmpCustRec(1).CUTOFFYN = fpCutOffYN.Text
  tmpCustRec(1).TAXEXPT = fpTaxExpt.Text
  tmpCustRec(1).SRCIT = fpSrCit.Text
  tmpCustRec(1).USEDRAFT = fpUseDraft.Text
  tmpCustRec(1).AcctType = QPTrim$(fpAcctType.Text)
  tmpCustRec(1).BankName = QPTrim$(fpBankName.Text)
  tmpCustRec(1).BANKLOC = QPTrim$(fpBankLoc.Text)
  tmpCustRec(1).TRANSIT = QPTrim$(fpTransit.Text)
  tmpCustRec(1).BankAcct = QPTrim$(fpBankAcct.Text)
  tmpCustRec(1).BILLCMNT = QPTrim$(fpBillCmnt.Text)
  tmpCustRec(1).PAYCMNT = QPTrim$(fpPayCmnt.Text)
  tmpCustRec(1).PumpCode = QPTrim$(fpPumpCode.Text)
  tmpCustRec(1).USERCODE1 = QPTrim$(fpUserCode1.Text)
  tmpCustRec(1).USERCODE2 = QPTrim$(fpUserCode2.Text)
  tmpCustRec(1).ProRatePCT = Val(QPTrim$(Str$(fpProRatePCT.Text)))
  tmpCustRec(1).HHMSG1 = QPTrim$(fpHHMsg1.Text)
  tmpCustRec(1).HHMSG2 = QPTrim$(fpHHMsg2.Text)
  tmpCustRec(1).HHMSG3 = QPTrim$(fpHHMsg3.Text)

  For cnt = 0 To 14
    tmpCustRec(1).serv(cnt + 1).Ratecode = QPTrim$(fpServCode(cnt).Text)
    tmpCustRec(1).serv(cnt + 1).RMtrType = QPTrim$(fpServMType(cnt).Text)
  Next
  For cnt = 0 To 3
    tmpCustRec(1).FlatRates(cnt + 1).FRDESC = QPTrim$(fpFlatDesc(cnt).Text)
    tmpCustRec(1).FlatRates(cnt + 1).FRAMT = Val(QPTrim$(Str$(fpFlatAmt(cnt).Text)))
    If fpFlatFreq(cnt).ListIndex = 0 Then
      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = "R"
    ElseIf fpFlatFreq(cnt).ListIndex = 1 Then
      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = "N"
    Else
      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = " "
    End If
    tmpCustRec(1).FlatRates(cnt + 1).REVSRC = Val(QPTrim$(Str$(fpFlatRevSrc(cnt).Text)))
    tmpCustRec(1).FlatRates(cnt + 1).NumMin = Val(QPTrim$(Str$(fpFlatMin(cnt).Text)))
  Next
  For cnt = 0 To 1
    tmpCustRec(1).Monthly(cnt + 1).AMTOWED = fpMonOwed(cnt)
    tmpCustRec(1).Monthly(cnt + 1).TotAmtPD = fpMonPaid(cnt)
    tmpCustRec(1).Monthly(cnt + 1).PayAmt = fpMonAmt(cnt)
    tmpCustRec(1).Monthly(cnt + 1).RevSource = fpMonRev(cnt)
  Next
  tmpCustRec(1).MFEE1 = fpMemFee(0)
  tmpCustRec(1).MFEE2 = fpMemFee(1)

  For cnt = 0 To 6
    tmpCustRec(1).LocMeters(cnt + 1).MtrNum = QPTrim$(fpMtrSerial(cnt))
    If Len(QPTrim$(fpMtrMulti(cnt).Text)) > 0 Then
      tmpCustRec(1).LocMeters(cnt + 1).MTRMulti = fpMtrMulti(cnt)
    Else
      tmpCustRec(1).LocMeters(cnt + 1).MTRMulti = -1
    End If
    tmpCustRec(1).LocMeters(cnt + 1).MTRType = QPTrim$(fpLocMType(cnt).Text)
    tmpCustRec(1).LocMeters(cnt + 1).MtrUnit = QPTrim$(fpLocUnit(cnt).Text)
    If Len(QPTrim$(fpMtrUser(cnt).Text)) > 0 Then
      tmpCustRec(1).LocMeters(cnt + 1).NumUser = fpMtrUser(cnt)
    Else
      tmpCustRec(1).LocMeters(cnt + 1).NumUser = -1
    End If
    tmpCustRec(1).LocMeters(cnt + 1).InsDate = Date2Num(fpLocMtrIns(cnt).Text)
    If Len(QPTrim$(fpLocMtrCur(cnt).Text)) > 0 Then
    'If Not Len(QPTrim$(fpLocMtrCur(cnt).Text)) < 0 Then
      tmpCustRec(1).LocMeters(cnt + 1).CurRead = fpLocMtrCur(cnt)
    End If
    If Len(QPTrim$(fpLocMtrPre(cnt).Text)) > 0 Then
    'If Len(QPTrim$(fpLocMtrPre(cnt).Text)) < 0 Then
      tmpCustRec(1).LocMeters(cnt + 1).PrevRead = fpLocMtrPre(cnt)
    End If
    tmpCustRec(1).LocMeters(cnt + 1).CurDate = Date2Num(fpLocMLRDate(cnt).Text)
    tmpCustRec(1).LocMeters(cnt + 1).MtrIDNO = QPTrim$(fpMtrIDNO(cnt).Text)
'put new field here
    If Not RecNo& > 0 Then
      tmpCustRec(1).LocMeters(cnt + 1).MtrLat = 0
      tmpCustRec(1).LocMeters(cnt + 1).MtrLng = 0
    End If
'
'put new field here
    'no no can't do the clear thing because of editing cust during meter read entry etc.
'    If ClearRFlag Then
'      tmpCustRec(1).LocMeters(cnt + 1).ReadFlag = "N"
'    Else
      tmpCustRec(1).LocMeters(cnt + 1).ReadFlag = tmpCustRec(2).LocMeters(cnt + 1).ReadFlag
'    End If
  Next
  'fpPrnBillYN.Text

'06/04/19 DW Deep Run   Clintwood
  PBYN = QPTrim(fpPrnBillYN.Text)
  If Len(PBYN) < 1 Then  'default to Yes
    PBYN = "Y"
  Else
    PBYN = UCase$(PBYN)
  End If
  tmpCustRec(1).PrnBillYN = PBYN
  
  tmpCustRec(1).FillPad = ""
  tmpCustRec(1).ChkByte = Chr$(5) 'changed this on 2/10/05 because of conversion
  
  DoEvents
  
  UBHandle = FreeFile
  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
  UBLog PWUser + (" Opened Custfile, SaveCust/AddEdit, for - " + Str(RecNo&) + " with " + Str(CustRecLen) + " len ")
  If RecNo& > 0 Then
    Put #UBHandle, RecNo&, tmpCustRec(1)
  Else
    RecNo& = (LOF(UBHandle) / CustRecLen) + 1
    Put #UBHandle, RecNo&, tmpCustRec(1)
  End If
  Close UBHandle
  UBLog PWUser + (" Closed Custfile, SaveCust/AddEdit ")
  UBLog PWUser + " Saved Acct: " + Str(RecNo&) + "," + QPTrim$(fpStatus.Text) + "," + QPTrim$(fpCustName.Text) + "," + QPTrim$(fpBook.Text) + "-" + QPTrim$(fpSeqNumb.Text)
  
  If UpDateOwner Then             'need to save new owner rec also
    Call UBSaveOwnerInfo(RecNo&)
  End If
  
  If RecNo& > 0 Then
    If tmpCustRec(1).SEARCH <> tmpCustRec(2).SEARCH Then
      ReindexFlag = True
    End If
    If tmpCustRec(1).CustName <> tmpCustRec(2).CustName Then
      ReindexFlag = True
    End If
    If (tmpCustRec(1).Book <> tmpCustRec(2).Book) Then
      ReindexFlag = True
    End If
    If (tmpCustRec(1).SEQNUMB <> tmpCustRec(2).SEQNUMB) Then
      ReindexFlag = True
    End If
    For cnt = 1 To 7
      If tmpCustRec(1).LocMeters(cnt).CurRead <> tmpCustRec(2).LocMeters(cnt).CurRead Then
        UBLog PWUser + " Saved Acct: " + Str(RecNo&) + ",changed Curr read - " + Str(tmpCustRec(2).LocMeters(cnt).CurRead) + " to " + Str(tmpCustRec(1).LocMeters(cnt).CurRead)
      End If
      If tmpCustRec(1).LocMeters(cnt).PrevRead <> tmpCustRec(2).LocMeters(cnt).PrevRead Then
        UBLog PWUser + " Saved Acct: " + Str(RecNo&) + ",changed Prev read - " + Str(tmpCustRec(2).LocMeters(cnt).PrevRead) + " to " + Str(tmpCustRec(1).LocMeters(cnt).PrevRead)
      End If
    Next
  Else  'adding new account set flag to reindex
    ReindexFlag = True
  End If
  DoEvents
  If ReindexFlag Then
    ReIndexSystem False
    DoEvents
  End If
  Close
  Erase tmpCustRec
  BlockInput False
  Call UPDateOK
  ActivateControls frmCustAddEdit
  
End Sub

Private Sub LoadCustInfo2Form()
  Dim tmpCustRec As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  CustRecLen = Len(tmpCustRec)
  
  RecNo& = Val(frmCustAddEdit.fpCustRecNo)
  frmCustAddEdit.fpCustRecNo = 0
  UBLog PWUser + " Edit Acct: " + Str(RecNo&)
  UBHandle = FreeFile
  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
  UBLog PWUser + (" Open custfile on Loadcust2form for- " + Str(RecNo&) + " in custaddedit")
  Get #UBHandle, RecNo&, tmpCustRec
  Close UBHandle
  UBLog PWUser + (" Closed custfile on Loadcust2form in custaddedit")
  If CustHasMsg(RecNo) Then
    MsgAlertTimer.Enabled = True
    'MsgRec = tmpCustRec.MessageRec
  End If
  
  If tmpCustRec.LastTrans > 0 Then
    TransRec = tmpCustRec.LastTrans
  End If
  
  If tmpCustRec.Status = "F" Then
    FinalFlag = True
    fpBook.Enabled = False
    fpSeqNumb.Enabled = False
  Else
    FinalFlag = False
    fpBook.Enabled = True
    fpSeqNumb.Enabled = True
  End If
  
  OldBook$ = tmpCustRec.Book + "-" + tmpCustRec.SEQNUMB

  LabelAcctNo.Caption = RecNo&
  fpBook = tmpCustRec.Book
  fpSeqNumb = tmpCustRec.SEQNUMB
  fpStatus.Text = " " + tmpCustRec.Status
  fpOpenDate = Num2Date(tmpCustRec.OPENDATE)
  fpSearch = QPTrim$(tmpCustRec.SEARCH)
  fpCustName = QPTrim$(tmpCustRec.CustName)
  'LblInfo.Caption = QPTrim$(tmpCustRec.CustName)
  fpAddr1 = QPTrim$(tmpCustRec.ADDR1)
  fpAddr2 = QPTrim$(tmpCustRec.ADDR2)
  fpServAddr = QPTrim$(tmpCustRec.ServAddr)
  fpCity = QPTrim$(tmpCustRec.CITY)
  fpState = QPTrim$(tmpCustRec.STATE)
  fpZip = QPTrim$(tmpCustRec.ZIPCODE)
  fpDPCode = QPTrim$(tmpCustRec.DPCode)
  fpHPhone.Text = QPTrim$(tmpCustRec.HPHONE)
  fpWPhone = QPTrim$(tmpCustRec.WPHONE)
'Stop
'here
  fpGroupCde.col = 0
  fpGroupCde.SearchText = Str$(tmpCustRec.GroupCodeRec)
  fpGroupCde.Action = 0
  If fpGroupCde.SearchIndex <> -1 Then
    fpGroupCde.ListIndex = fpGroupCde.SearchIndex
  Else
    fpGroupCde.ListIndex = 0
  End If

  fpSoSec = QPTrim$(tmpCustRec.SOSEC)
  fpDrvLic = QPTrim$(tmpCustRec.DRVLIC)
  fpCustType = QPTrim$(tmpCustRec.CUSTTYPE)
  fpAddr911 = QPTrim$(tmpCustRec.Addr911)
  If QPTrim$(tmpCustRec.BillTo) = "O" Then
    fpBillTo.ListIndex = 1
  Else
    fpBillTo.ListIndex = 0
  End If
  fpBillCopy = QPTrim$(Str$(tmpCustRec.BILLCOPY))
  fpPostRte = QPTrim$(tmpCustRec.POSTRTE)
  If tmpCustRec.BILLCYCL >= 0 Then
    fpBillCycl = QPTrim$(Str$(tmpCustRec.BILLCYCL))
  Else
    fpBillCycl = ""
  End If
  fpZone = QPTrim$(tmpCustRec.ZONE)
  If tmpCustRec.Seq >= 0 Then
    fpSeq = QPTrim$(Str$(tmpCustRec.Seq))
  Else
    fpSeq = ""
  End If
  Select Case tmpCustRec.CASHONLY
  Case "N", " "
    fpCashOnly.Value = ValueFalse
  Case Else
    fpCashOnly.Value = ValueTrue
  End Select
  Select Case tmpCustRec.LATEFEE
  Case "N", " "
    fpLateFee.Value = ValueFalse
  Case Else
    fpLateFee.Value = ValueTrue
  End Select
  Select Case tmpCustRec.CUTOFFYN
  Case "N", " "
    fpCutOffYN.Value = ValueFalse
  Case Else
    fpCutOffYN.Value = ValueTrue
  End Select
  Select Case tmpCustRec.TAXEXPT
  Case "N", " "
    fpTaxExpt.Value = ValueFalse
  Case Else
    fpTaxExpt.Value = ValueTrue
  End Select
  Select Case tmpCustRec.SRCIT
  Case "N", " "
    fpSrCit.Value = ValueFalse
  Case Else
    fpSrCit.Value = ValueTrue
  End Select
  Select Case tmpCustRec.USEDRAFT
  Case "Y"
    fpUseDraft.Value = ValueTrue
  Case Else
    fpUseDraft.Value = ValueFalse
  End Select
  
  fpAcctType = QPTrim$(tmpCustRec.AcctType)
  fpBankName = QPTrim$(tmpCustRec.BankName)
  fpBankLoc = QPTrim$(tmpCustRec.BANKLOC)
  fpTransit = QPTrim$(tmpCustRec.TRANSIT)
  fpBankAcct = QPTrim$(tmpCustRec.BankAcct)
  fpBillCmnt = QPTrim$(tmpCustRec.BILLCMNT)
  fpPayCmnt = QPTrim$(tmpCustRec.PAYCMNT)
  fpPumpCode = QPTrim$(tmpCustRec.PumpCode)
  fpUserCode1 = QPTrim$(tmpCustRec.USERCODE1)
  fpUserCode2 = QPTrim$(tmpCustRec.USERCODE2)
  fpProRatePCT = QPTrim$(Str$(tmpCustRec.ProRatePCT))
  fpHHMsg1 = QPTrim$(tmpCustRec.HHMSG1)
  fpHHMsg2 = QPTrim$(tmpCustRec.HHMSG2)
  fpHHMsg3 = QPTrim$(tmpCustRec.HHMSG3)
  
  For cnt = 0 To 14
    fpServCode(cnt).Text = QPTrim$(tmpCustRec.serv(cnt + 1).Ratecode)
    fpServMType(cnt).Text = QPTrim$(tmpCustRec.serv(cnt + 1).RMtrType)
  Next
  
  For cnt = 0 To 3
    fpFlatDesc(cnt).Text = QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRDESC)
    fpFlatAmt(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).FRAMT))
    If QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ) = "R" Then
      fpFlatFreq(cnt).ListIndex = 0
    ElseIf QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ) = "N" Then
      fpFlatFreq(cnt).ListIndex = 1
    Else
      fpFlatFreq(cnt).ListIndex = -1
    End If
    fpFlatRevSrc(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).REVSRC))
    fpFlatMin(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).NumMin))
  Next
  
  For cnt = 0 To 1
    'fpMonOwed(Cnt).Text = QPTrim$(Str$(tmpCustRec.Monthly(Cnt + 1).AMTOWED))
    fpMonOwed(cnt) = tmpCustRec.Monthly(cnt + 1).AMTOWED
    fpMonPaid(cnt) = tmpCustRec.Monthly(cnt + 1).TotAmtPD
    fpMonAmt(cnt) = tmpCustRec.Monthly(cnt + 1).PayAmt
    fpMonRev(cnt) = tmpCustRec.Monthly(cnt + 1).RevSource
  Next
  fpMemFee(0) = tmpCustRec.MFEE1
  fpMemFee(1) = tmpCustRec.MFEE2
  
  For cnt = 0 To 6
    fpMtrSerial(cnt) = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrNum)
    If tmpCustRec.LocMeters(cnt + 1).MTRMulti >= 0 Then
      fpMtrMulti(cnt) = tmpCustRec.LocMeters(cnt + 1).MTRMulti
    End If
    fpLocMType(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MTRType)
    fpLocUnit(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrUnit)
    If tmpCustRec.LocMeters(cnt + 1).NumUser > 0 Then
      fpMtrUser(cnt) = tmpCustRec.LocMeters(cnt + 1).NumUser
    End If
    fpLocMtrIns(cnt).Text = Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate)
    'If tmpCustRec.LocMeters(cnt + 1).CurRead > 0 Then
    If tmpCustRec.LocMeters(cnt + 1).CurRead >= 0 Then
      fpLocMtrCur(cnt) = tmpCustRec.LocMeters(cnt + 1).CurRead
    End If
    'If tmpCustRec.LocMeters(cnt + 1).PrevRead > 0 Then
    If tmpCustRec.LocMeters(cnt + 1).PrevRead >= 0 Then
      fpLocMtrPre(cnt) = tmpCustRec.LocMeters(cnt + 1).PrevRead
    End If
    fpLocMLRDate(cnt).Text = Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate)
    fpMtrIDNO(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrIDNO)
  Next
  fpPrnBillYN.Text = tmpCustRec.PrnBillYN
  DoEvents
End Sub
    
Private Function CheckSaveCustFile%()
  Dim Changed As Boolean
  Dim chkCustRec As NewUBCustRecType
  Dim UBHandle As Integer, Enoughtosave As Boolean
  Dim CustRecLen As Integer
  CustRecLen = Len(chkCustRec)
  Enoughtosave = True
  If UpDateOwner Then 'check owner info
    Changed = True
    GoTo DoneCustChk
  End If
  
  If RecNo& > 0 Then
    UBHandle = FreeFile
    Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
    UBLog PWUser + (" Open custfile on Checksavecustfile in custaddedit for cust " + Str(RecNo&))
    Get #UBHandle, RecNo&, chkCustRec
    Close UBHandle
    UBLog PWUser + (" closed custfile on Checksavecustfile in custaddedit for cust " + Str(RecNo&))
    If QPTrim$(chkCustRec.Book) <> QPTrim$(fpBook.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.SEQNUMB) <> QPTrim$(fpSeqNumb.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.Status) <> QPTrim$(fpStatus.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If chkCustRec.OPENDATE <> Date2Num(fpOpenDate.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.SEARCH) <> QPTrim$(fpSearch.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.CustName) <> QPTrim$(fpCustName.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.ADDR1) <> QPTrim$(fpAddr1.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.ADDR2) <> QPTrim$(fpAddr2.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.ServAddr) <> QPTrim$(fpServAddr.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.CITY) <> QPTrim$(fpCity.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.STATE) <> QPTrim$(fpState.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Val(chkCustRec.ZIPCODE) <> Val(fpZip.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.DPCode) <> QPTrim$(fpDPCode.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPStripStuff$(chkCustRec.HPHONE) <> QPStripStuff$(fpHPhone.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPStripStuff$(chkCustRec.WPHONE) <> QPStripStuff$(fpWPhone.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Val(chkCustRec.SOSEC) <> Val(fpSoSec.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.DRVLIC) <> QPTrim$(fpDrvLic.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.CUSTTYPE) <> QPTrim$(fpCustType.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.Addr911) <> QPTrim$(fpAddr911.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    fpGroupCde.col = 0
    If chkCustRec.GroupCodeRec <> Val(fpGroupCde.ColText) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Len(QPTrim$(chkCustRec.BillTo)) > 0 Then
    If QPTrim$(chkCustRec.BillTo) <> Mid$(fpBillTo.Text, 1, 1) Then
      Changed = True
      GoTo DoneCustChk
    End If
    End If
    If chkCustRec.BILLCOPY <> Val(fpBillCopy) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.POSTRTE) <> QPTrim$(fpPostRte.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Len(QPTrim$(fpBillCycl.Text)) = 0 Then
      If Not chkCustRec.BILLCYCL = -32767 Then
        Changed = True
        GoTo DoneCustChk
      End If
    Else
      If chkCustRec.BILLCYCL <> Val(fpBillCycl) Then
        Changed = True
        GoTo DoneCustChk
      End If
    End If
    If QPTrim$(chkCustRec.ZONE) <> QPTrim$(fpZone.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Len(QPTrim$(fpSeq.Text)) = 0 Then
      If Not chkCustRec.Seq = -32767 Then
        Changed = True
        GoTo DoneCustChk
      End If
    Else
      If chkCustRec.Seq <> Val(fpSeq) Then
        Changed = True
        GoTo DoneCustChk
      End If
    End If
    If QPTrim$(chkCustRec.CASHONLY) <> fpCashOnly.Text Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.LATEFEE) <> fpLateFee.Text Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.CUTOFFYN) <> fpCutOffYN.Text Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.TAXEXPT) <> fpTaxExpt.Text Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.SRCIT) <> fpSrCit.Text Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Len(QPTrim$(chkCustRec.USEDRAFT)) > 0 Then
      If QPTrim$(chkCustRec.USEDRAFT) <> fpUseDraft.Text Then
        Changed = True
        GoTo DoneCustChk
      End If
    End If
    If QPTrim$(chkCustRec.AcctType) <> QPTrim$(fpAcctType.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.BankName) <> QPTrim$(fpBankName.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.BANKLOC) <> QPTrim$(fpBankLoc.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.TRANSIT) <> QPTrim$(fpTransit.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.BankAcct) <> QPTrim$(fpBankAcct.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.BILLCMNT) <> QPTrim$(fpBillCmnt.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.PAYCMNT) <> QPTrim$(fpPayCmnt.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.PumpCode) <> QPTrim$(fpPumpCode.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.USERCODE1) <> QPTrim$(fpUserCode1.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.USERCODE2) <> QPTrim$(fpUserCode2.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If chkCustRec.ProRatePCT <> fpProRatePCT Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.HHMSG1) <> QPTrim$(fpHHMsg1.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.HHMSG2) <> QPTrim$(fpHHMsg2.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    If QPTrim$(chkCustRec.HHMSG3) <> QPTrim$(fpHHMsg3.Text) Then
      Changed = True
      GoTo DoneCustChk
    End If
    For cnt = 0 To 14
      If QPTrim$(chkCustRec.serv(cnt + 1).Ratecode) <> QPTrim$(fpServCode(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If QPTrim$(chkCustRec.serv(cnt + 1).RMtrType) <> QPTrim$(fpServMType(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
    Next
    
    For cnt = 0 To 3
      If QPTrim$(chkCustRec.FlatRates(cnt + 1).FRDESC) <> QPTrim$(fpFlatDesc(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.FlatRates(cnt + 1).FRAMT <> fpFlatAmt(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If QPTrim$(chkCustRec.FlatRates(cnt + 1).FRFREQ) <> Mid$(fpFlatFreq(cnt).ColText, 1, 1) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.FlatRates(cnt + 1).REVSRC <> fpFlatRevSrc(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.FlatRates(cnt + 1).NumMin <> fpFlatMin(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
    Next
    
    For cnt = 0 To 1
      If chkCustRec.Monthly(cnt + 1).AMTOWED <> fpMonOwed(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.Monthly(cnt + 1).TotAmtPD <> fpMonPaid(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.Monthly(cnt + 1).PayAmt <> fpMonAmt(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If chkCustRec.Monthly(cnt + 1).RevSource <> fpMonRev(cnt) Then
        Changed = True
        GoTo DoneCustChk
      End If
    Next
    If chkCustRec.MFEE1 <> fpMemFee(0) Then
      Changed = True
      GoTo DoneCustChk
    End If
    
    If chkCustRec.MFEE2 <> fpMemFee(1) Then
      Changed = True
      GoTo DoneCustChk
    End If
    
    For cnt = 0 To 6
      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrNum) <> QPTrim$(fpMtrSerial(cnt)) Then
        Changed = True
        GoTo DoneCustChk
      End If
'NOTE: DO NOT change this comparsion. Must be done this way to maintain
'      compatibility with old way of storing a blank numeric field. Old
'      method stored the maximum negitive value of the numeric variable type
'      (i.e. integer, double, long etc.) to represent a blank field. Since
'      the a meter multiplier can not be a negitive value, I am storing a
'      -1 (negitive one) to represent this in the new version.
      
      If chkCustRec.LocMeters(cnt + 1).MTRMulti <= 0 Then
        If Val(fpMtrMulti(cnt)) > 0 Then
          Changed = True
          GoTo DoneCustChk
        End If
      ElseIf chkCustRec.LocMeters(cnt + 1).MTRMulti <> Val(fpMtrMulti(cnt)) Then
        Changed = True
        GoTo DoneCustChk
      End If
            
      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MTRType) <> QPTrim$(fpLocMType(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      
      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrUnit) <> QPTrim$(fpLocUnit(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If Len(QPTrim$(fpMtrUser(cnt).Text)) > 0 Then
        If chkCustRec.LocMeters(cnt + 1).NumUser <> fpMtrUser(cnt) Then
          Changed = True
          GoTo DoneCustChk
        End If
      Else
        If chkCustRec.LocMeters(cnt + 1).NumUser <> -1 Then
          Changed = True
          GoTo DoneCustChk
        End If
      End If
      
      If chkCustRec.LocMeters(cnt + 1).InsDate <> Date2Num(fpLocMtrIns(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If Len(QPTrim$(fpLocMtrCur(cnt).Text)) > 0 Then
        If chkCustRec.LocMeters(cnt + 1).CurRead <> fpLocMtrCur(cnt) Then
         Changed = True
         GoTo DoneCustChk
        End If
      End If
      If Len(QPTrim$(fpLocMtrPre(cnt).Text)) > 0 Then
        If chkCustRec.LocMeters(cnt + 1).PrevRead <> fpLocMtrPre(cnt) Then
          Changed = True
          GoTo DoneCustChk
        End If
      End If
      If chkCustRec.LocMeters(cnt + 1).CurDate <> Date2Num(fpLocMLRDate(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrIDNO) <> QPTrim$(fpMtrIDNO(cnt).Text) Then
        Changed = True
        GoTo DoneCustChk
      End If
    Next
  Else
'    If fpstatus.ListIndex = -1 Then
'      changed = true
'      GoTo DoneCustChk
'    End If
    If Len(QPTrim$(fpSearch.Text)) > 0 Then
      Changed = True
      GoTo DoneCustChk
    End If
    If Len(QPTrim$(fpCustName.Text)) > 0 Then
      Changed = True
      GoTo DoneCustChk
    End If
  End If
  
DoneCustChk:
  DoEvents

  If Changed Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      CheckSaveCustFile% = False
    Case True
      CheckSaveCustFile% = True
    Case 1
      CheckSaveCustFile% = 1
    End Select
  Else
    CheckSaveCustFile% = False
  End If
  DoEvents
End Function

Private Sub ChkFormatBookSeqN()
  Dim TBook As String
  Dim TSeqN As String
  TBook = QPTrim$(Me.fpBook)
  TSeqN = QPTrim$(Me.fpSeqNumb)
  Me.fpBook = FmtBook$(Me.fpBook)
  Me.fpSeqNumb = FmtSeqN$(Me.fpSeqNumb)
End Sub

Private Function ChkCustInfoOK%()
Dim Enoughtosave As Boolean
  Enoughtosave = True
  ChkCustInfoOK = False   'assume the worst.
  If RecNo& > 0 Then
 ' If Not FinalFlag Then     'if this account isn't in final
    
    Call ChkFormatBookSeqN
    NBook$ = Me.fpBook + "-" + Me.fpSeqNumb
    If QPTrim$(fpStatus.Text) = "F" And (OldBook$ <> NBook$) Then
        vaTabPro1.ActiveTab = 0
        DoEvents
        MsgBox "   Final Status Does NOT Allow Location #'s!   " + Chr$(13) + Chr$(13) + "   Please enter a new Status or Location ", vbOKOnly, "ERROR!"
        Me.fpStatus.SetFocus
        ChkCustInfoOK = False
    Else
    If (OldBook$ <> NBook$) And (NBook$ <> "00-000000") Then
      'If fpStatus.Text = "A" Or fpStatus.Text = "P" Then
      If Not Val(Me.fpBook) > 0 And Val(Me.fpSeqNumb) > 0 Then
        vaTabPro1.ActiveTab = 0
        DoEvents
        MsgBox "   Invalid Book!   " + Chr$(13) + Chr$(13) + "   Please enter a new Book number   ", vbOKOnly, "ERROR!"
        Me.fpBook.SetFocus
        ChkCustInfoOK = False
      Else
      'if they changed the book-seq list num
      If Chk4DupeLocation(Me.fpBook, Me.fpSeqNumb) Then
        If Len(OldBook$) > 1 Then
          Me.fpBook = Left$(OldBook$, 2)
          Me.fpSeqNumb = Mid$(OldBook$, 4)
        Else
          Me.fpBook = ""
          Me.fpSeqNumb = ""
        End If
        vaTabPro1.ActiveTab = 0
        DoEvents
        MsgBox "   Duplicate Location Number Found!   " + Chr$(13) + Chr$(13) + "   Please enter a new location number   ", vbOKOnly, "ERROR!"
        If Me.fpBook.Enabled = True Then
          Me.fpBook.SetFocus
        Else
          Me.fpStatus.SetFocus
        End If
        ChkCustInfoOK = False
      Else
        ChkCustInfoOK = True
      End If
      End If
    Else
      ChkCustInfoOK = True
    End If
   End If
'  Else
'    ChkCustInfoOK = True
'  End If
  Else
    If Chk4DupeLocation(Me.fpBook, Me.fpSeqNumb) Then
      If Len(OldBook$) > 1 Then
        Me.fpBook = Left$(OldBook$, 2)
        Me.fpSeqNumb = Mid$(OldBook$, 4)
      Else
        Me.fpBook = ""
        Me.fpSeqNumb = ""
      End If
      vaTabPro1.ActiveTab = 0
      DoEvents
      MsgBox "   Duplicate Location Number Found!   " + Chr$(13) + Chr$(13) + "   Please enter a new location number   ", vbOKOnly, "ERROR!"
      If fpBook.Enabled = True Then
        Me.fpBook.SetFocus
      Else
        Me.fpStatus.SetFocus
      End If
      ChkCustInfoOK = False
    Else
      ChkCustInfoOK = True
    End If

    If fpStatus.ListIndex = -1 Then
      Enoughtosave = False
      GoTo DoneChk
    End If
    If Not Len(QPTrim$(fpSearch.Text)) > 0 Then
      If QPTrim$(fpStatus.Text) = "A" Then
        Enoughtosave = False
        GoTo DoneChk
      End If
    End If
    If Not Len(QPTrim$(fpCustName.Text)) > 0 Then
      Enoughtosave = False
      GoTo DoneChk
    End If
  End If
Exit Function
DoneChk:
  DoEvents
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If Enoughtosave = False Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "The Name, Search Name and Status"
    MsgText(3) = "are Required Fields."
    MsgText(4) = ""
    MsgText(5) = "Please Enter This Information."
    GetOKorNot MsgText(), True
    ChkCustInfoOK = False
  Else
    ChkCustInfoOK = True
  End If

End Function

Private Function Chk4DupeLocation(Book$, SeqNum$)
  Dim TBookSeq  As Long, NumBookSeq As Long
  Dim BookSeqLen As Integer, Handle As Integer
  Dim DupeFlag As Boolean
  ReDim UBBookSeq(1) As BookSeqRecType
  Chk4DupeLocation = False    'assume it's ok
  TBookSeq = Val(Book$ + SeqNum$)
  BookSeqLen = Len(UBBookSeq(1))
  If FileSize(UBPath$ + "UBOOKSEQ.DAT") > 0 Then
    Handle = FreeFile
    Open UBPath$ + "UBOOKSEQ.DAT" For Random Shared As Handle Len = BookSeqLen
    NumBookSeq = LOF(Handle) \ BookSeqLen
    For CntL = 1 To NumBookSeq
      Get Handle, CntL, UBBookSeq(1)
      If UBBookSeq(1).BookSeq = TBookSeq& Then
        If Not QPTrim$(fpStatus.Text) = "A" Or QPTrim$(fpStatus.Text) = "P" Then
       ' If Not fpstatus.Text = "A" Or fpstatus.Text = "P" Then
          If TBookSeq& <= 0 Then
            Exit For
          End If
        End If
        DupeFlag = True
        Exit For
      End If
    Next
  End If
  Close Handle
  If DupeFlag Then
    Chk4DupeLocation = True
  End If
  'Erase UBBookSeq
End Function

Private Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = fpCmdMsg.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      fpCmdMsg.ForeColor = &H80000012
      fpCmdMsg.FontSize = BtnFnt
    Case 2
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 0.7
    Case 3
      fpCmdMsg.ForeColor = &H80000011
      fpCmdMsg.FontSize = BtnFnt - 1.4
    Case 4
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.1
    Case 5
      fpCmdMsg.ForeColor = &H80000010
      fpCmdMsg.FontSize = BtnFnt - 2.8
    Case 6
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 3.5
    Case 7
      fpCmdMsg.ForeColor = &H8000000F
      fpCmdMsg.FontSize = BtnFnt - 4.2
    Case 8
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 4.9
    Case 9
      fpCmdMsg.ForeColor = &H8000000E
      fpCmdMsg.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
'  DoEvents
End Sub

Private Sub CustAddEdLoadRateCodes()
  Dim RateRec As UBRateTblRecType
  Dim Handle As Integer, cnt As Integer, zz As Integer
  Dim RateRecLen As Integer, NumOfRates As Integer
  Dim tmp As String
  
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  NumOfRates = GetNumRateRecs%
  For zz = 0 To MaxRevsCnt - 1  'Set the revenue captions
    Me.PG3RevLBL(zz).Caption = QPTrim$(UBSetUpRec(1).Revenues(zz + 1).RevName)
    Me.fpServCode(zz).AddItem " " + Chr9 + "NO RATE"
  Next
'  If UBSetUpRec(1).HHDEVICE = "E" Then
'    fpcmdMtrCoordinates.Enabled = True
'  Else
'    fpcmdMtrCoordinates.Enabled = False
'  End If
  RateRecLen = Len(RateRec)
  Handle = FreeFile
  Open UBPath + "ubrate.dat" For Random Shared As Handle Len = RateRecLen
  For cnt = 1 To NumOfRates
    Get #Handle, cnt, RateRec
    tmp$ = QPTrim$(RateRec.Ratecode) + Chr9 + QPTrim$(RateRec.RATEDESC)
    For zz = 0 To MaxRevsCnt - 1
      Me.fpServCode(zz).AddItem tmp$
    Next
  Next
  Close
  'this will insure that the end user dosn't select a code or
  'select a disable code/meter on non rate code services
  For zz = 0 To MaxRevsCnt - 1
    Me.fpServCode(zz).ListIndex = 0
    If UBSetUpRec(1).Revenues(zz + 1).USERATE <> "Y" Then
      Me.fpServCode(zz).Enabled = False
    End If
    If UBSetUpRec(1).Revenues(zz + 1).UseMtr <> "Y" Then
      Me.fpServMType(zz).Enabled = False
    End If
  Next
End Sub
Private Sub WorkOrders()
  DeActivateControls Me
  frmWorkOrderEntry.fpCmdConHist.Visible = False
  frmWorkOrderEntry.fpCmdMsg.Visible = False
  frmWorkOrderEntry.fpCmdOwner.Visible = False
  frmWorkOrderEntry.fpCmdTranHist.Visible = False
  frmWorkOrderEntry.fpCustRecNo = RecNo&
  frmWorkOrderEntry.Wheretogo frmCustAddEdit, frmWorkOrderEntry, , 55
  'Load frmWorkOrderEntry
  frmWorkOrderEntry.Show
  DoEvents
End Sub
Private Sub setup4new()
  RecNo = 0
  BeenDone = False
  TransRec = 0
  fpCustRecNo = 0
  NBook$ = ""
  MsgRec = 0
  OldBook = ""
  FinalFlag = False
  UpDateOwner = False

  fpBook = ""
  fpSeqNumb = ""
  fpSearch = ""
  fpCustName = ""
  fpAddr1 = ""
  fpAddr2 = ""
  fpServAddr = ""
  fpDPCode = ""
  fpHPhone.Text = ""
  fpWPhone = ""
  fpSoSec = ""
  fpDrvLic = ""
  fpCustType = ""
  fpAddr911 = ""
  fpPostRte = ""
  fpBillCycl = ""
  fpZone = ""
  fpSeq = ""
  fpCashOnly.Value = ValueFalse
  fpLateFee.Value = ValueTrue
  fpCutOffYN.Value = ValueTrue
  fpTaxExpt.Value = ValueFalse
  fpSrCit.Value = ValueFalse
  fpUseDraft.Value = ValueFalse
  fpAcctType = ""
  fpBankName = ""
  fpBankLoc = ""
  fpTransit = ""
  fpBankAcct = ""
  fpBillCmnt = ""
  fpPayCmnt = ""
  fpPumpCode = ""
  fpUserCode1 = ""
  fpUserCode2 = ""
  fpProRatePCT = 100
  fpHHMsg1 = ""
  fpHHMsg2 = ""
  fpHHMsg3 = ""
  For cnt = 0 To 14
    fpServCode(cnt).Text = ""
    fpServMType(cnt).Text = ""
  Next
  
  For cnt = 0 To 3
    fpFlatDesc(cnt).Text = ""
    fpFlatAmt(cnt).Text = ""
    fpFlatFreq(cnt).ListIndex = -1
    fpFlatRevSrc(cnt).Text = ""
    fpFlatMin(cnt).Text = ""
  Next
  
  For cnt = 0 To 1
    fpMonOwed(cnt) = 0
    fpMonPaid(cnt) = 0
    fpMonAmt(cnt) = 0
    fpMonRev(cnt) = ""
  Next
  fpMemFee(0) = 0
  fpMemFee(1) = 0
  
  For cnt = 0 To 6
    fpMtrSerial(cnt) = ""
    fpLocMType(cnt).Text = ""
    fpLocUnit(cnt).Text = ""
    fpLocMtrIns(cnt).Text = Format(Now, "mm/dd/yyyy")
    fpLocMtrCur(cnt) = 0
    fpLocMtrPre(cnt) = 0
    fpLocMLRDate(cnt).Text = Format(Now, "mm/dd/yyyy")
    fpMtrIDNO(cnt).Text = ""
  Next
   vaTabPro1.ActiveTab = 0
   'fpBook.SetFocus
End Sub

