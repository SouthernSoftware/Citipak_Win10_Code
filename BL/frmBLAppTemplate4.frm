VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLAppTemplate4 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Application Renewal Template #4"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11760
   Icon            =   "frmBLAppTemplate4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   8652
      Left            =   1056
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   48
      Width           =   8052
      _Version        =   196609
      _ExtentX        =   14203
      _ExtentY        =   15261
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
      TabCount        =   2
      Orientation     =   1
      ThreeD          =   -1  'True
      ActiveTabBold   =   -1  'True
      GrayAreaColor   =   13684944
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      EarMarkColorDark=   13684944
      BookRingColor   =   13684944
      BookShowMetalSpine=   -1  'True
      PageEarMarkColorDark=   13684944
      DataFormat      =   ""
      AutoSizeChildren=   2
      BookCornerGuardWidth=   90
      BookCornerGuardLength=   375
      ThreeDOuterLight=   13684944
      DataField       =   ""
      TabCaption      =   "frmBLAppTemplate4.frx":08CA
      PageEarMarkPictureNext=   "frmBLAppTemplate4.frx":0A7E
      PageEarMarkPicturePrev=   "frmBLAppTemplate4.frx":0A9A
      EarMarkPictureNext=   "frmBLAppTemplate4.frx":0AB6
      EarMarkPicturePrev=   "frmBLAppTemplate4.frx":0AD2
      Begin ImpproLib.vaImprint vaImprint1 
         Height          =   8460
         Left            =   15
         TabIndex        =   41
         Top             =   45
         Width           =   7650
         _Version        =   196609
         _ExtentX        =   13494
         _ExtentY        =   14922
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
         FrameThreeDStyle=   3
         Picture         =   "frmBLAppTemplate4.frx":0AEE
         Begin LpLib.fpCombo fpcmbStartMonth 
            Height          =   315
            Left            =   2070
            TabIndex        =   12
            Tag             =   "From this drop down box select the month the license renewal becomes valid."
            Top             =   4230
            Width           =   1170
            _Version        =   196608
            _ExtentX        =   2064
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":0B0A
         End
         Begin LpLib.fpCombo fpcmbPenMonth 
            Height          =   315
            Left            =   1350
            TabIndex        =   16
            Tag             =   "From this drop down box select the month after which the license renewal becomes invalid."
            Top             =   4515
            Width           =   1170
            _Version        =   196608
            _ExtentX        =   2064
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":0E71
         End
         Begin LpLib.fpCombo fpcmbPenDay 
            Height          =   315
            Left            =   2490
            TabIndex        =   17
            Tag             =   "From this drop down box select the last day after which license renewal becomes invalid."
            Top             =   4515
            Width           =   585
            _Version        =   196608
            _ExtentX        =   1032
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":11D8
         End
         Begin LpLib.fpCombo fpcmbStartDay 
            Height          =   315
            Left            =   3210
            TabIndex        =   13
            Tag             =   "From this drop down box select the first day the license renewal becomes valid."
            Top             =   4230
            Width           =   570
            _Version        =   196608
            _ExtentX        =   1005
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":153F
         End
         Begin LpLib.fpCombo fpcmbDiscMonth 
            Height          =   300
            Left            =   2490
            TabIndex        =   4
            Tag             =   "From this drop down box select the month the information for this license renewal should be returned."
            Top             =   2115
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":18A6
         End
         Begin LpLib.fpCombo fpcmbDiscDay 
            Height          =   300
            Left            =   3120
            TabIndex        =   5
            Tag             =   "From this drop down box select the last day the information for this license renewal should be returned."
            Top             =   2115
            Width           =   570
            _Version        =   196608
            _ExtentX        =   1005
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":1C0D
         End
         Begin LpLib.fpCombo fpcmbYear1 
            Height          =   300
            Left            =   3795
            TabIndex        =   1
            Tag             =   $"frmBLAppTemplate4.frx":1F74
            Top             =   765
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":228B
         End
         Begin LpLib.fpCombo fpcmbYear2 
            Height          =   300
            Left            =   3840
            TabIndex        =   6
            Tag             =   $"frmBLAppTemplate4.frx":25F2
            Top             =   2115
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":28BB
         End
         Begin LpLib.fpCombo fpcmbYear3 
            Height          =   300
            Left            =   3840
            TabIndex        =   14
            Tag             =   $"frmBLAppTemplate4.frx":2C22
            Top             =   4230
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":2EEB
         End
         Begin LpLib.fpCombo fpcmbYear4 
            Height          =   300
            Left            =   6240
            TabIndex        =   15
            Tag             =   $"frmBLAppTemplate4.frx":3252
            Top             =   4230
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":351B
         End
         Begin LpLib.fpCombo fpcmbYear5 
            Height          =   300
            Left            =   3165
            TabIndex        =   18
            Tag             =   $"frmBLAppTemplate4.frx":3882
            Top             =   4515
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   529
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.5
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
            ColDesigner     =   "frmBLAppTemplate4.frx":3B4B
         End
         Begin EditLib.fpText fptxtTownOf 
            Height          =   252
            Left            =   2448
            TabIndex        =   0
            Tag             =   $"frmBLAppTemplate4.frx":3EB2
            Top             =   336
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
         Begin EditLib.fpText fptxtAdd 
            Height          =   252
            Left            =   2352
            TabIndex        =   8
            Tag             =   "Enter the town's mailing address in this field."
            Top             =   3456
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
            Height          =   252
            Left            =   2064
            TabIndex        =   9
            Tag             =   "Enter the town's mailing name here."
            Top             =   3696
            Width           =   1836
            _Version        =   196608
            _ExtentX        =   3238
            _ExtentY        =   444
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
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
         Begin EditLib.fpText fptxtState 
            Height          =   252
            Left            =   3888
            TabIndex        =   10
            Tag             =   "Enter the town's state in this field (ex. NC = North Carolina)."
            Top             =   3696
            Width           =   300
            _Version        =   196608
            _ExtentX        =   529
            _ExtentY        =   444
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
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
         Begin EditLib.fpMask fptxtZip 
            Height          =   252
            Left            =   4176
            TabIndex        =   11
            Tag             =   "Enter the town's postal code here."
            Top             =   3696
            Width           =   876
            _Version        =   196608
            _ExtentX        =   1545
            _ExtentY        =   444
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
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
         Begin EditLib.fpText fptxtOrdinance 
            Height          =   252
            Left            =   3792
            TabIndex        =   2
            Tag             =   $"frmBLAppTemplate4.frx":3F93
            Top             =   1632
            Width           =   2124
            _Version        =   196608
            _ExtentX        =   3746
            _ExtentY        =   444
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
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
         Begin EditLib.fpDateTime fptxtAdoptDate 
            Height          =   252
            Left            =   432
            TabIndex        =   3
            Tag             =   "Enter the date here which the town ordinance was adopted."
            Top             =   1872
            Width           =   1116
            _Version        =   196608
            _ExtentX        =   1968
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
            Text            =   "6/11/2003"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
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
            PopUpType       =   0
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
            StartMonth      =   3
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fptxtTownBody 
            Height          =   252
            Left            =   384
            TabIndex        =   7
            Tag             =   $"frmBLAppTemplate4.frx":4079
            Top             =   2928
            Width           =   1788
            _Version        =   196608
            _ExtentX        =   3154
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
         Begin VB.Label Label91 
            BackColor       =   &H80000009&
            Caption         =   ")"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   6768
            TabIndex        =   141
            Top             =   4272
            Width           =   108
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label90 
            BackColor       =   &H80000009&
            Caption         =   ","
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3792
            TabIndex        =   140
            Top             =   4320
            Width           =   108
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label78 
            BackColor       =   &H80000009&
            Caption         =   ","
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3744
            TabIndex        =   135
            Top             =   2160
            Width           =   108
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000009&
            Caption         =   "PAGE 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5760
            TabIndex        =   133
            Top             =   816
            Width           =   636
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblOrdCity 
            BackColor       =   &H80000009&
            Caption         =   "TOWN NAME"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   768
            TabIndex        =   131
            Top             =   1680
            Width           =   1404
         End
         Begin VB.Label lblRespectTown 
            BackColor       =   &H80000009&
            Caption         =   "Your Town"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   432
            TabIndex        =   130
            Top             =   2688
            Width           =   2892
         End
         Begin VB.Label Label40 
            BackColor       =   &H80000009&
            Caption         =   "RECEIPTS FOR EACH CLASSIFICATION THAT APPLIES TO YOUR BUSINESS."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   81
            Top             =   7392
            Width           =   6300
         End
         Begin VB.Label Label39 
            BackColor       =   &H80000009&
            Caption         =   "NOT RESULT IN ANY ADDITIONAL COST TO BUSINESSES. PLEASE REPORT GROSS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   80
            Top             =   7200
            Width           =   6828
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000009&
            Caption         =   ". THIS WILL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   5952
            TabIndex        =   79
            Top             =   7008
            Width           =   1020
         End
         Begin VB.Label lblTownOrd 
            BackColor       =   &H80000009&
            Caption         =   "TOWN ORDINANCE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3456
            TabIndex        =   78
            Top             =   7008
            Width           =   2508
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Application for Town Licenses"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   2400
            TabIndex        =   77
            Top             =   3936
            Width           =   2364
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label37 
            BackColor       =   &H80000009&
            Caption         =   "_________________________"
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
            Left            =   1104
            TabIndex        =   76
            Top             =   6288
            Width           =   2556
         End
         Begin VB.Label Label36 
            BackColor       =   &H80000009&
            Caption         =   "_________________________"
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
            Left            =   1104
            TabIndex        =   75
            Top             =   6048
            Width           =   2556
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000009&
            Caption         =   "____________________________"
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
            Left            =   4368
            TabIndex        =   74
            Top             =   5856
            Width           =   2460
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000009&
            Caption         =   "_________________________"
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
            Left            =   1104
            TabIndex        =   73
            Top             =   6528
            Width           =   2172
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000009&
            Caption         =   "__________________________"
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
            Left            =   4560
            TabIndex        =   72
            Top             =   6528
            Width           =   2220
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000009&
            Caption         =   "911:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   432
            TabIndex        =   71
            Top             =   6000
            Width           =   396
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000009&
            Caption         =   "MAIL:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3888
            TabIndex        =   70
            Top             =   5520
            Width           =   540
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            Caption         =   "TRADING AS:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   69
            Top             =   5040
            Width           =   1752
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000009&
            Caption         =   ", "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3072
            TabIndex        =   68
            Top             =   4608
            Width           =   108
         End
         Begin VB.Label Label34 
            BackColor       =   &H80000009&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   4368
            TabIndex        =   67
            Top             =   2160
            Width           =   36
         End
         Begin VB.Label Label31 
            BackColor       =   &H80000009&
            Caption         =   "adopted"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   5952
            TabIndex        =   66
            Top             =   1680
            Width           =   684
         End
         Begin VB.Label Label44 
            BackColor       =   &H80000009&
            Caption         =   "PHONE: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   432
            TabIndex        =   65
            Top             =   6480
            Width           =   684
         End
         Begin VB.Label Label43 
            BackColor       =   &H80000009&
            Caption         =   "(or start of business in "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   4416
            TabIndex        =   64
            Top             =   4272
            Width           =   1788
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000009&
            Caption         =   "____________________________"
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
            Left            =   4368
            TabIndex        =   63
            Top             =   5664
            Width           =   2556
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000009&
            Caption         =   "NAME OF APPLICANT:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   432
            TabIndex        =   62
            Top             =   4848
            Width           =   2136
         End
         Begin VB.Label Label32 
            BackColor       =   &H80000009&
            Caption         =   "---------------------------------------------------------------------------------------------------------"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   168
            Left            =   432
            TabIndex        =   61
            Top             =   3072
            Width           =   6612
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000009&
            Caption         =   "Respectfully,"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   432
            TabIndex        =   60
            Top             =   2448
            Width           =   1056
         End
         Begin VB.Label Label21 
            BackColor       =   &H80000009&
            Caption         =   "and "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   432
            TabIndex        =   59
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000009&
            Caption         =   "For Year: "
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
            Left            =   3072
            TabIndex        =   58
            Top             =   816
            Width           =   684
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000009&
            Caption         =   "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
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
            Left            =   1584
            TabIndex        =   57
            Top             =   576
            Width           =   4044
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000009&
            Caption         =   "please complete and return this form with the required "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   1584
            TabIndex        =   56
            Top             =   1920
            Width           =   4308
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000009&
            Caption         =   "Town Ordinance #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2256
            TabIndex        =   55
            Top             =   1680
            Width           =   1464
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000009&
            Caption         =   " License (BPOL) Tax promulgated by Virginia Code Section 58.1-3700 et seq."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   54
            Top             =   1440
            Width           =   6396
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000009&
            Caption         =   "       For the purpose of computing Business, Professional and Occupational"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   432
            TabIndex        =   53
            Top             =   1200
            Width           =   6300
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            Caption         =   "Dear Business Owner:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   456
            TabIndex        =   52
            Top             =   960
            Width           =   1800
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000009&
            Caption         =   " information no later than"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   432
            TabIndex        =   51
            Top             =   2160
            Width           =   2052
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000009&
            Caption         =   "For period beginning"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   432
            TabIndex        =   50
            Top             =   4272
            Width           =   1836
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000009&
            Caption         =   "and ending"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   432
            TabIndex        =   49
            Top             =   4560
            Width           =   900
         End
         Begin VB.Label Label10 
            BackColor       =   &H80000009&
            Caption         =   "BUSINESS ADDRESS:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   432
            TabIndex        =   48
            Top             =   5328
            Width           =   1740
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000009&
            Caption         =   "MAIL:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   432
            TabIndex        =   47
            Top             =   5520
            Width           =   540
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000009&
            Caption         =   "PHONE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3888
            TabIndex        =   46
            Top             =   6480
            Width           =   732
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000009&
            Caption         =   "HOME ADDRESS:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3888
            TabIndex        =   45
            Top             =   5328
            Width           =   1452
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000009&
            Caption         =   "A SEPARATE LICENSE WILL BE ISSUED FOR EACH TYPE OF BUSINESS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   44
            Top             =   6816
            Width           =   5676
         End
         Begin VB.Label Label28 
            BackColor       =   &H80000009&
            Caption         =   "PERFORMED, AS REQUIRED PER THE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   43
            Top             =   7008
            Width           =   2988
         End
         Begin VB.Label lblTownName2 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Town Of"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   2352
            TabIndex        =   42
            Top             =   3216
            Width           =   2412
         End
      End
      Begin ImpproLib.vaImprint vaImprint2 
         Height          =   8520
         Left            =   -22725
         TabIndex        =   82
         Top             =   -23565
         Width           =   7680
         _Version        =   196609
         _ExtentX        =   13547
         _ExtentY        =   15028
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
         Enabled         =   0   'False
         BackColor       =   -2147483639
         Caption         =   ""
         FrameThreeDStyle=   3
         Picture         =   "frmBLAppTemplate4.frx":410D
         Begin LpLib.fpCombo fpcmbLicRetDay 
            Height          =   315
            Left            =   6390
            TabIndex        =   39
            Tag             =   "Select a day of the month from the drop down list that is the final day the business license fee can be paid without penalty."
            Top             =   6390
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":4129
         End
         Begin LpLib.fpCombo fpcmbLicRetMonth 
            Height          =   315
            Left            =   5130
            TabIndex        =   38
            Tag             =   "Select the month from the drop down list that is the last valid month the business license fee can be paid."
            Top             =   6390
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":4490
         End
         Begin LpLib.fpCombo fpcmbAppRetDay 
            Height          =   315
            Left            =   5520
            TabIndex        =   37
            Tag             =   $"frmBLAppTemplate4.frx":47F7
            Top             =   6090
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":4883
         End
         Begin LpLib.fpCombo fpcmbAppRetMonth 
            Height          =   315
            Left            =   4275
            TabIndex        =   36
            Tag             =   $"frmBLAppTemplate4.frx":4BEA
            Top             =   6090
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":4C8B
         End
         Begin LpLib.fpCombo fpcmbRepairDay 
            Height          =   315
            Left            =   2955
            TabIndex        =   33
            Tag             =   $"frmBLAppTemplate4.frx":4FF2
            Top             =   3555
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":5086
         End
         Begin LpLib.fpCombo fpcmbRepairMonth 
            Height          =   315
            Left            =   2325
            TabIndex        =   32
            Tag             =   $"frmBLAppTemplate4.frx":53ED
            Top             =   3555
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":5483
         End
         Begin LpLib.fpCombo fpcmbContDay 
            Height          =   315
            Left            =   2955
            TabIndex        =   30
            Tag             =   "From this drop down list select the last day to include for total receipts for last year for Contracting."
            Top             =   2970
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":57EA
         End
         Begin LpLib.fpCombo fpcmbContMonth 
            Height          =   315
            Left            =   2325
            TabIndex        =   29
            Tag             =   "From this drop down list select the last month to include for total receipts for last year for Contracting."
            Top             =   2970
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":5B51
         End
         Begin LpLib.fpCombo fpcmbFinDay 
            Height          =   315
            Left            =   2955
            TabIndex        =   26
            Tag             =   $"frmBLAppTemplate4.frx":5EB8
            Top             =   2400
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":5F52
         End
         Begin LpLib.fpCombo fpcmbFinMonth 
            Height          =   315
            Left            =   2325
            TabIndex        =   25
            Tag             =   $"frmBLAppTemplate4.frx":62B9
            Top             =   2400
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":6351
         End
         Begin LpLib.fpCombo fpcmbRetailDay 
            Height          =   315
            Left            =   2955
            TabIndex        =   23
            Tag             =   "From this drop down list select the last day to include for total receipts for last year for Retail Merchants."
            Top             =   1770
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":66B8
         End
         Begin LpLib.fpCombo fpcmbRetailMonth 
            Height          =   315
            Left            =   2325
            TabIndex        =   22
            Tag             =   "From this drop down list select the last month to include for total receipts for last year for Retail Merchants."
            Top             =   1770
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":6A1F
         End
         Begin LpLib.fpCombo fpcmbWholeDay 
            Height          =   315
            Left            =   2955
            TabIndex        =   20
            Tag             =   "From this drop down list select the last day to include for total receipts for last year for Wholesale Merchants."
            Top             =   1155
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":6D86
         End
         Begin LpLib.fpCombo fpcmbWholeMonth 
            Height          =   315
            Left            =   2325
            TabIndex        =   19
            Tag             =   "From this drop down list select the last month to include for total receipts for last year for Wholesale Merchants."
            Top             =   1155
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":70ED
         End
         Begin LpLib.fpCombo fpcmbYear6 
            Height          =   315
            Left            =   3600
            TabIndex        =   21
            Tag             =   $"frmBLAppTemplate4.frx":7454
            Top             =   1155
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":771D
         End
         Begin LpLib.fpCombo fpcmbYear7 
            Height          =   315
            Left            =   3600
            TabIndex        =   24
            Tag             =   $"frmBLAppTemplate4.frx":7A84
            Top             =   1770
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":7D4D
         End
         Begin LpLib.fpCombo fpcmbYear8 
            Height          =   315
            Left            =   3600
            TabIndex        =   27
            Tag             =   $"frmBLAppTemplate4.frx":80B4
            Top             =   2400
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":837D
         End
         Begin LpLib.fpCombo fpcmbYear9 
            Height          =   315
            Left            =   3600
            TabIndex        =   31
            Tag             =   $"frmBLAppTemplate4.frx":86E4
            Top             =   2970
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":89AD
         End
         Begin LpLib.fpCombo fpcmbYear10 
            Height          =   315
            Left            =   3600
            TabIndex        =   34
            Tag             =   $"frmBLAppTemplate4.frx":8D14
            Top             =   3555
            Width           =   540
            _Version        =   196608
            _ExtentX        =   952
            _ExtentY        =   556
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            ColDesigner     =   "frmBLAppTemplate4.frx":8FDD
         End
         Begin EditLib.fpMask fptxtPhone 
            Height          =   252
            Left            =   480
            TabIndex        =   35
            Tag             =   "Enter the telephone number in this field where business license related inquiries should be directed."
            Top             =   4320
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
         Begin VB.Label Label33 
            BackColor       =   &H80000009&
            Caption         =   "PAGE 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   6288
            TabIndex        =   134
            Top             =   816
            Width           =   636
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label30 
            BackColor       =   &H80000009&
            Caption         =   "TO AVOID PENALTY."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   132
            Top             =   6480
            Width           =   1788
         End
         Begin VB.Label Label89 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2856
            TabIndex        =   129
            Top             =   1200
            Width           =   84
         End
         Begin VB.Label Label88 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   3504
            TabIndex        =   128
            Top             =   1200
            Width           =   84
         End
         Begin VB.Label Label87 
            BackColor       =   &H80000009&
            Caption         =   " as shown by applicants records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4152
            TabIndex        =   127
            Top             =   1200
            Width           =   2556
         End
         Begin VB.Label Label86 
            BackColor       =   &H80000009&
            Caption         =   "LICENSE FEES ARE DUE PRIOR TO "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   2304
            TabIndex        =   126
            Top             =   6480
            Width           =   2796
         End
         Begin VB.Label Label85 
            BackColor       =   &H80000009&
            Caption         =   "APPLICATION MUST BE RETURNED PRIOR TO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   125
            Top             =   6144
            Width           =   3804
         End
         Begin VB.Label Label84 
            BackColor       =   &H80000009&
            Caption         =   "for assistance."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1776
            TabIndex        =   124
            Top             =   4320
            Width           =   1452
         End
         Begin VB.Label Label83 
            BackColor       =   &H80000009&
            Caption         =   "WHOLESALE MERCHANT:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   123
            Top             =   960
            Width           =   2040
         End
         Begin VB.Label Label82 
            BackColor       =   &H80000009&
            Caption         =   "Gross Receipts through"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   504
            TabIndex        =   122
            Top             =   1200
            Width           =   1884
         End
         Begin VB.Label Label81 
            BackColor       =   &H80000009&
            Caption         =   "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
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
            Left            =   1728
            TabIndex        =   121
            Top             =   576
            Width           =   4044
         End
         Begin VB.Label Label80 
            BackColor       =   &H80000009&
            Caption         =   "For Year: 20XX"
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
            Left            =   3216
            TabIndex        =   120
            Top             =   768
            Width           =   1260
         End
         Begin VB.Label Label79 
            BackColor       =   &H80000009&
            Caption         =   "Signature ________________________"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            TabIndex        =   119
            Top             =   5184
            Width           =   3180
         End
         Begin VB.Label Label77 
            BackColor       =   &H80000009&
            Caption         =   "***IMPORTANT***"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   2688
            TabIndex        =   118
            Top             =   5856
            Width           =   1692
         End
         Begin VB.Label Label76 
            BackColor       =   &H80000009&
            Caption         =   "Print Name_________________________"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3648
            TabIndex        =   117
            Top             =   5520
            Width           =   3468
         End
         Begin VB.Label Label75 
            BackColor       =   &H80000009&
            Caption         =   "$ __________________"
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
            Left            =   5256
            TabIndex        =   116
            Top             =   1440
            Width           =   1788
         End
         Begin VB.Label Label74 
            BackColor       =   &H80000009&
            Caption         =   "If uncertain of your business classification(s), please call the Town Office at"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   115
            Top             =   4080
            Width           =   5964
         End
         Begin VB.Label Label73 
            BackColor       =   &H80000009&
            Caption         =   "OF EACH YEAR TO AVOID PENALTY AND INTEREST. INTENTIONALLY PROVIDING"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   114
            Top             =   6816
            Width           =   6732
         End
         Begin VB.Label Label72 
            BackColor       =   &H80000009&
            Caption         =   "INSUFFICIENT OR INACCURATE INFORMATION MAY RESULT IN LEGAL RECOURSE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   113
            Top             =   7152
            Width           =   6828
         End
         Begin VB.Label Label71 
            BackColor       =   &H80000009&
            Caption         =   "BY THE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   480
            TabIndex        =   112
            Top             =   7488
            Width           =   636
         End
         Begin VB.Label Label70 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2856
            TabIndex        =   111
            Top             =   1824
            Width           =   84
         End
         Begin VB.Label Label69 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   3528
            TabIndex        =   110
            Top             =   1824
            Width           =   84
         End
         Begin VB.Label Label68 
            BackColor       =   &H80000009&
            Caption         =   "as shown by applicants records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4200
            TabIndex        =   109
            Top             =   1824
            Width           =   2460
         End
         Begin VB.Label Label67 
            BackColor       =   &H80000009&
            Caption         =   "RETAIL MERCHANT:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   108
            Top             =   1584
            Width           =   2040
         End
         Begin VB.Label Label66 
            BackColor       =   &H80000009&
            Caption         =   "Gross Receipts through"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   504
            TabIndex        =   107
            Top             =   1824
            Width           =   1884
         End
         Begin VB.Label Label65 
            BackColor       =   &H80000009&
            Caption         =   "$ __________________"
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
            Left            =   5256
            TabIndex        =   106
            Top             =   2064
            Width           =   1788
         End
         Begin VB.Label Label64 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2856
            TabIndex        =   105
            Top             =   2448
            Width           =   84
         End
         Begin VB.Label Label63 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   3528
            TabIndex        =   104
            Top             =   2448
            Width           =   84
         End
         Begin VB.Label Label62 
            BackColor       =   &H80000009&
            Caption         =   "as shown by applicants records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4200
            TabIndex        =   103
            Top             =   2448
            Width           =   2460
         End
         Begin VB.Label Label61 
            BackColor       =   &H80000009&
            Caption         =   "FINANCIAL, REAL ESTATE AND PROFESSIONAL:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   102
            Top             =   2208
            Width           =   4632
         End
         Begin VB.Label Label60 
            BackColor       =   &H80000009&
            Caption         =   "Gross Receipts through"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   504
            TabIndex        =   101
            Top             =   2448
            Width           =   1884
         End
         Begin VB.Label Label59 
            BackColor       =   &H80000009&
            Caption         =   "$ __________________"
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
            Left            =   5256
            TabIndex        =   100
            Top             =   2688
            Width           =   1788
         End
         Begin VB.Label Label58 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2856
            TabIndex        =   99
            Top             =   3024
            Width           =   84
         End
         Begin VB.Label Label57 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   3528
            TabIndex        =   98
            Top             =   3024
            Width           =   84
         End
         Begin VB.Label Label56 
            BackColor       =   &H80000009&
            Caption         =   "as shown by applicants records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4224
            TabIndex        =   97
            Top             =   3024
            Width           =   2460
         End
         Begin VB.Label Label55 
            BackColor       =   &H80000009&
            Caption         =   "CONTRACTING:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   96
            Top             =   2784
            Width           =   2040
         End
         Begin VB.Label Label54 
            BackColor       =   &H80000009&
            Caption         =   "Gross Receipts through"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   504
            TabIndex        =   95
            Top             =   3024
            Width           =   1884
         End
         Begin VB.Label Label53 
            BackColor       =   &H80000009&
            Caption         =   "$ __________________"
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
            Left            =   5256
            TabIndex        =   94
            Top             =   3264
            Width           =   1788
         End
         Begin VB.Label Label52 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   2856
            TabIndex        =   93
            Top             =   3600
            Width           =   84
         End
         Begin VB.Label Label51 
            BackColor       =   &H80000009&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   3528
            TabIndex        =   92
            Top             =   3600
            Width           =   84
         End
         Begin VB.Label Label49 
            BackColor       =   &H80000009&
            Caption         =   "as shown by applicants records"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4248
            TabIndex        =   91
            Top             =   3600
            Width           =   2460
         End
         Begin VB.Label Label48 
            BackColor       =   &H80000009&
            Caption         =   "REPAIR, PERSONAL or BUSINESS SERVICES:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   480
            TabIndex        =   90
            Top             =   3360
            Width           =   4104
         End
         Begin VB.Label Label47 
            BackColor       =   &H80000009&
            Caption         =   "Gross Receipts through"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   504
            TabIndex        =   89
            Top             =   3600
            Width           =   1884
         End
         Begin VB.Label Label46 
            BackColor       =   &H80000009&
            Caption         =   "$ __________________"
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
            Left            =   5256
            TabIndex        =   88
            Top             =   3840
            Width           =   1788
         End
         Begin VB.Label Label45 
            BackColor       =   &H80000009&
            Caption         =   "I do affirm that the foregoing figures are true, complete and accurate to the best of my knowledge."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   480
            TabIndex        =   87
            Top             =   4656
            Width           =   6012
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label42 
            BackColor       =   &H80000009&
            Caption         =   "OF EACH YEAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   6096
            TabIndex        =   86
            Top             =   6144
            Width           =   1260
         End
         Begin VB.Label lblPenTownname 
            BackColor       =   &H80000009&
            Caption         =   "YOUR TOWN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1152
            TabIndex        =   85
            Top             =   7488
            Width           =   2748
         End
         Begin VB.Label Label41 
            BackColor       =   &H80000009&
            Caption         =   "AS SET FORTH BY VIRGINIA CODE."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   3984
            TabIndex        =   84
            Top             =   7488
            Width           =   2988
         End
         Begin VB.Label lblPage2TownHead 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "YOURTOWN"
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
            Left            =   1632
            TabIndex        =   83
            Top             =   384
            Width           =   4044
         End
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   9570
      TabIndex        =   136
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to close this screen and return to the Town Setup screen."
      Top             =   6420
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
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
      ButtonDesigner  =   "frmBLAppTemplate4.frx":9344
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   690
      Left            =   9570
      TabIndex        =   137
      TabStop         =   0   'False
      Tag             =   "Press this 'Next App' button to close this application screen and open up the screen for application #5."
      Top             =   4530
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
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
      ButtonDesigner  =   "frmBLAppTemplate4.frx":9522
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   9570
      TabIndex        =   138
      TabStop         =   0   'False
      Tag             =   "Press 'Save' to save the currently active application as application #4. All fields will be committed to memory."
      Top             =   7365
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
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
      ButtonDesigner  =   "frmBLAppTemplate4.frx":9701
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   675
      Left            =   9570
      TabIndex        =   139
      TabStop         =   0   'False
      Tag             =   "Press this 'Last App' to close this screen and open the screen for application #3."
      Top             =   5490
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
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
      ButtonDesigner  =   "frmBLAppTemplate4.frx":98DD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   9555
      TabIndex        =   142
      Tag             =   $"frmBLAppTemplate4.frx":9ABC
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   3360
      Width           =   1875
      _Version        =   131072
      _ExtentX        =   3307
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmBLAppTemplate4.frx":9B86
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   10080
      TabIndex        =   143
      Top             =   1152
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   876
      Left            =   9360
      Top             =   3156
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
      Left            =   9456
      TabIndex        =   144
      Top             =   4128
      Width           =   2052
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   972
      Left            =   9492
      Top             =   1764
      Width           =   1980
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Application #4 Virginia"
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
      Height          =   876
      Left            =   9600
      TabIndex        =   28
      Top             =   1824
      Width           =   1740
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPntScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBLAppTemplate4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmBLAppTemplate4
  frmBLTownSetup.fpcmbAppType.SetFocus
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    frmBLMessageBoxJr.Label1.Caption = "This application is intended for use only in the State of Virginia. There is a reference to a Virginia Code in the final paragraph of this application."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    frmBLMessageBoxJr.Label1.Top = 450
    frmBLMessageBoxJr.Label1.Height = 1300
    frmBLMessageBoxJr.Label1.Caption = "Some of the discretionary values initially appearing on this page are supplied from the Town Setup screen. If other application templates have been used then some of the values here may have carried over from them. PLEASE REVIEW ALL values to make sure they reflect the CURRENT situation."
    frmBLMessageBoxJr.Show vbModal
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    fptxtTownOf.ToolTipText = ""
    fpcmbYear1.ToolTipText = ""
    fptxtOrdinance.ToolTipText = ""
    fptxtAdoptDate.ToolTipText = ""
    fpcmbDiscMonth.ToolTipText = ""
    fpcmbDiscDay.ToolTipText = ""
    fpcmbYear2.ToolTipText = ""
    fptxtTownBody.ToolTipText = ""
    fptxtAdd.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    fpcmbStartMonth.ToolTipText = ""
    fpcmbStartDay.ToolTipText = ""
    fpcmbYear3.ToolTipText = ""
    fpcmbYear4.ToolTipText = ""
    fpcmbPenMonth.ToolTipText = ""
    fpcmbPenDay.ToolTipText = ""
    fpcmbYear5.ToolTipText = ""
    fpcmbWholeMonth.ToolTipText = ""
    fpcmbWholeDay.ToolTipText = ""
    fpcmbYear6.ToolTipText = ""
    fpcmbRetailMonth.ToolTipText = ""
    fpcmbRetailDay.ToolTipText = ""
    fpcmbYear7.ToolTipText = ""
    fpcmbFinMonth.ToolTipText = ""
    fpcmbFinDay.ToolTipText = ""
    fpcmbYear8.ToolTipText = ""
    fpcmbContMonth.ToolTipText = ""
    fpcmbContDay.ToolTipText = ""
    fpcmbYear9.ToolTipText = ""
    fpcmbRepairMonth.ToolTipText = ""
    fpcmbRepairDay.ToolTipText = ""
    fpcmbYear10.ToolTipText = ""
    fptxtPhone.ToolTipText = ""
    fpcmbAppRetMonth.ToolTipText = ""
    fpcmbAppRetDay.ToolTipText = ""
    fpcmbLicRetMonth.ToolTipText = ""
    fpcmbLicRetDay.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdNext.ToolTipText = ""
    cmdLast.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "Press this button to activate/deactivate instructional balloons."
'    fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'    fpcmbYear1.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtOrdinance.ToolTipText = "Enter the town ordinance pertaining to business license renewals (ex. #103 Uniform B.P.O.L. Policy) here."
'    fptxtAdoptDate.ToolTipText = "Enter the day this town ordinance was adopted."
'    fpcmbDiscMonth.ToolTipText = "Select the month the information for this license renewal should be returned."
'    fpcmbDiscDay.ToolTipText = "Select the first day which the new business license will be valid."
'    fpcmbYear2.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtTownBody.ToolTipText = "Enter the town's official governing body (Town Council) here."
'    fptxtAdd.ToolTipText = "Enter your town's mailing address here."
'    fptxtCity.ToolTipText = "Enter your town's mailing name here."
'    fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
'    fptxtZip.ToolTipText = "Enter your town's zip code here."
'    fpcmbStartMonth.ToolTipText = "Select the first month which the new business license will be valid."
'    fpcmbStartDay.ToolTipText = "Select the first day which the new business license will be valid."
'    fpcmbYear3.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbYear4.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbPenMonth.ToolTipText = "Select the last month which the new business license will be valid."
'    fpcmbPenDay.ToolTipText = "Select the last day which the new business license will be valid."
'    fpcmbYear5.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbWholeMonth.ToolTipText = "Select the last month to include for total receipts for last year for Wholesale Merchants."
'    fpcmbWholeDay.ToolTipText = "Select the last day to include for total receipts for last year for Wholesale Merchants."
'    fpcmbYear6.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbRetailMonth.ToolTipText = "Select the last month to include for total receipts for last year for Retail."
'    fpcmbRetailDay.ToolTipText = "Select the last day to include for total receipts for last year for Retail."
'    fpcmbYear7.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbFinMonth.ToolTipText = "Select the last month to include for total receipts for last year for Financial, Real Estate and Professional services."
'    fpcmbFinDay.ToolTipText = "Select the last day to include for total receipts for last year for Financial, Real Estate and Professional services."
'    fpcmbYear8.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbContMonth.ToolTipText = "Select the last month to include for total receipts for last year for Contracting."
'    fpcmbContDay.ToolTipText = "Select the last day to include for total receipts for last year for Contracting."
'    fpcmbYear9.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fpcmbRepairMonth.ToolTipText = "Select the last month to include for total receipts for last year for Repair, Personal or Business Services."
'    fpcmbRepairDay.ToolTipText = "Select the last day to include for total receipts for last year for Repair, Personal or Business Services."
'    fpcmbYear10.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtPhone.ToolTipText = "Enter the town's official phone number here."
'    fpcmbAppRetMonth.ToolTipText = "Select the month by which Business License renewal applications must be returned."
'    fpcmbAppRetDay.ToolTipText = "Select the day by which Business License renewal applications must be returned."
'    fpcmbLicRetMonth.ToolTipText = "Select the month by which Business License fees must be paid."
'    fpcmbLicRetDay.ToolTipText = "Select the day by which Business License fees must be paid."
'    cmdNext.ToolTipText = "Press to move to application template #5."
'    cmdLast.ToolTipText = "Press to move to business application #3."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = "Press to save the data on this screen."
  End If
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim TempCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fptxtAdd.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid mailing address for your town."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtAdd.BackColor = &H80FFFF
    fptxtAdd.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtState.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's state."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtState.BackColor = &H80FFFF
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's mailing name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtCity.BackColor = &H80FFFF
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtTownOf.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's official name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtTownOf.BackColor = &H80FFFF
    fptxtTownOf.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtTownBody.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town body's official name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtTownBody.BackColor = &H80FFFF
    fptxtTownBody.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtZip.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's zip code."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtZip.BackColor = &H80FFFF
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdoptDate.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town ordinance adoption date."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtAdoptDate.BackColor = &H80FFFF
    fptxtAdoptDate.SetFocus
    Exit Sub
  End If

  If QPTrim$(fptxtOrdinance.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the BPOL town ordinance."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtOrdinance.BackColor = &H80FFFF
    fptxtOrdinance.SetFocus
    Exit Sub
  End If

  If QPTrim$(fptxtPhone.Text) = "(" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the appropriate telephone number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    vaTabPro1.ActiveTab = 0
    fptxtPhone.BackColor = &H80FFFF
    fptxtPhone.SetFocus
    Exit Sub
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
      TownRec.AppTownOf = QPTrim(fptxtTownOf.Text)
      TownRec.AppAdd1 = QPTrim(fptxtAdd.Text)
      TownRec.AppCity = QPTrim(fptxtCity.Text)
      TownRec.AppState = QPTrim(fptxtState.Text)
      TownRec.AppZip = QPTrim(fptxtZip.Text)
      TownRec.AppPhone = QPTrim(fptxtPhone.Text)
      TownRec.AppMayorCouncil = QPTrim$(fptxtTownBody.Text)
      TownRec.AppWholeMonth = fpcmbWholeMonth.Text
      TownRec.AppWholeDay = fpcmbWholeDay.Text
      TownRec.AppRetailMonth = fpcmbRetailMonth.Text
      TownRec.AppRetailDay = fpcmbRetailDay.Text
      TownRec.AppFinMonth = fpcmbFinMonth.Text
      TownRec.AppFinDay = fpcmbFinDay.Text
      TownRec.AppContMonth = fpcmbContMonth.Text
      TownRec.AppContDay = fpcmbContDay.Text
      TownRec.AppRepairMonth = fpcmbRepairMonth.Text
      TownRec.AppRepairDay = fpcmbRepairDay.Text
      TownRec.AppPenMonth = QPTrim$(fpcmbPenMonth.Text)
      TownRec.AppPenDay = fpcmbPenDay.Text
      TownRec.AppFiscMonth = QPTrim$(fpcmbAppRetMonth.Text)
      TownRec.AppFiscDay = fpcmbAppRetDay.Text
      TownRec.AppStartMonth = QPTrim$(fpcmbStartMonth.Text)
      TownRec.AppStartDay = fpcmbStartDay.Text
      TownRec.AppLicRetMonth = QPTrim$(fpcmbLicRetMonth.Text)
      TownRec.AppLicRetDay = fpcmbLicRetDay.Text
      TownRec.AppDiscMonth = fpcmbDiscMonth.Text
      TownRec.AppDiscDay = fpcmbDiscDay.Text
      TownRec.AppAdoptDate = Date2Num(fptxtAdoptDate.Text)
      TownRec.AppCityOrd = QPTrim$(fptxtOrdinance.Text)
      TownRec.AppYrUpDown(1) = QPTrim$(fpcmbYear1.Text)
      TownRec.AppYrUpDown(2) = QPTrim$(fpcmbYear2.Text)
      TownRec.AppYrUpDown(3) = QPTrim$(fpcmbYear3.Text)
      TownRec.AppYrUpDown(4) = QPTrim$(fpcmbYear4.Text)
      TownRec.AppYrUpDown(5) = QPTrim$(fpcmbYear5.Text)
      TownRec.AppYrUpDown(6) = QPTrim$(fpcmbYear6.Text)
      TownRec.AppYrUpDown(7) = QPTrim$(fpcmbYear7.Text)
      TownRec.AppYrUpDown(8) = QPTrim$(fpcmbYear8.Text)
      TownRec.AppYrUpDown(9) = QPTrim$(fpcmbYear9.Text)
      TownRec.AppYrUpDown(10) = QPTrim$(fpcmbYear10.Text)
      TownRec.AppForm = 4
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
    TownRec.AppForm = 4
    TownRec.DLQNotice = 0
    TownRec.AppAdd1 = QPTrim(fptxtAdd.Text) '12
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
    TownRec.AppState = QPTrim$(fptxtState.Text)
    TownRec.AppCity = QPTrim$(fptxtCity.Text)
    TownRec.AppTownOf = QPTrim$(fptxtTownOf.Text)
    TownRec.AppZip = QPTrim$(fptxtZip.Text) '30
    TownRec.AppPct = 0
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppPhone = fptxtPhone.Text
    TownRec.AppDiscPct = 0
    TownRec.AppDiscMonth = fpcmbDiscMonth.Text
    TownRec.AppDiscDay = fpcmbDiscDay.Text
    TownRec.AppPenMonth = QPTrim$(fpcmbPenMonth.Text)
    TownRec.AppPenDay = CInt(fpcmbPenDay.Text)
    TownRec.AppFiscMonth = QPTrim$(fpcmbAppRetMonth.Text)
    TownRec.AppFiscDay = fpcmbAppRetDay.Text
    TownRec.AppMayorCouncil = QPTrim$(fptxtTownBody.Text)
    TownRec.AppWholeMonth = fpcmbWholeMonth.Text
    TownRec.AppWholeDay = fpcmbWholeDay.Text
    TownRec.AppRetailMonth = fpcmbRetailMonth.Text
    TownRec.AppRetailDay = fpcmbRetailDay.Text
    TownRec.AppFinMonth = fpcmbFinMonth.Text
    TownRec.AppFinDay = fpcmbFinDay.Text
    TownRec.AppContMonth = fpcmbContMonth.Text
    TownRec.AppContDay = fpcmbContDay.Text
    TownRec.AppRepairMonth = fpcmbRepairMonth.Text
    TownRec.AppRepairDay = fpcmbRepairDay.Text
    TownRec.AppStartMonth = QPTrim$(fpcmbStartMonth.Text)
    TownRec.AppStartDay = fpcmbStartDay.Text '61
    TownRec.AppLicRetMonth = QPTrim$(fpcmbLicRetMonth.Text)
    TownRec.AppLicRetDay = fpcmbLicRetDay.Text
    TownRec.AppAdoptDate = Date2Num(fptxtAdoptDate.Text)
    TownRec.AppPayBy = 0
    TownRec.AppCityOrd = QPTrim$(fptxtOrdinance.Text)
    TownRec.AppYrUpDown(1) = fpcmbYear1.Text
    TownRec.AppYrUpDown(2) = QPTrim$(fpcmbYear2.Text)
    TownRec.AppYrUpDown(3) = QPTrim$(fpcmbYear3.Text)
    TownRec.AppYrUpDown(4) = QPTrim$(fpcmbYear4.Text)
    TownRec.AppYrUpDown(5) = QPTrim$(fpcmbYear5.Text)
    TownRec.AppYrUpDown(6) = QPTrim$(fpcmbYear6.Text)
    TownRec.AppYrUpDown(7) = QPTrim$(fpcmbYear7.Text)
    TownRec.AppYrUpDown(8) = QPTrim$(fpcmbYear8.Text)
    TownRec.AppYrUpDown(9) = QPTrim$(fpcmbYear9.Text)
    TownRec.AppYrUpDown(10) = QPTrim$(fpcmbYear10.Text)
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = ""
    TownRec.DlqAdminTitle = ""
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = "" '71
    TownRec.DlqFax = ""
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = "" '81
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

  'added as a precaution to prevent the user from running application
  'renewal form #4 then coming here to save different data and then
  'trying to run application renewal reprints which will use this
  'latest saved data while the originals have the old data...now the
  'user will have to print applications over
  If Exist("artmpcus.dat") Then
    OpenTempCustRec TempHandle
    TempCnt = LOF(TempHandle) / Len(TempCustRec)
    If TempCnt > 0 Then
      Get TempHandle, 1, TempCustRec
      Close TempHandle
      If TempCustRec.AppType = 4 Then
        KillFile "artmpcus.dat"
      End If
    Else
      Close TempHandle
    End If
  End If
  
  frmBLSucSave.Label1.Caption = "Your renewal application notice #4 data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  Call cmdExit_Click
  frmBLTownSetup.fpcmbAppType.Text = "4. APP FORM C"
  frmBLTownSetup.fpcmdApps.Text = "F3 S&how App Type 4"
  
  MainLog ("Application #4 saved.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate4", "cmdSave_Click", Erl)
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
  DoEvents
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
    Case vbKeyF3:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%L"
      Call cmdLast_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLAppTemplate4.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  lblBalloon.Visible = False
'  fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'  fpcmbYear1.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtOrdinance.ToolTipText = "Enter the town ordinance pertaining to business license renewals (ex. #103 Uniform B.P.O.L. Policy) here."
'  fptxtAdoptDate.ToolTipText = "Enter the day this town ordinance was adopted."
'  fpcmbDiscMonth.ToolTipText = "Select the month the information for this license renewal should be returned."
'  fpcmbDiscDay.ToolTipText = "Select the first day which the new business license will be valid."
'  fpcmbYear2.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtTownBody.ToolTipText = "Enter the town's official governing body (Town Council) here."
'  fptxtAdd.ToolTipText = "Enter your town's mailing address here."
'  fptxtCity.ToolTipText = "Enter your town's mailing name here."
'  fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
'  fptxtZip.ToolTipText = "Enter your town's zip code here."
'  fpcmbStartMonth.ToolTipText = "Select the first month which the new business license will be valid."
'  fpcmbStartDay.ToolTipText = "Select the first day which the new business license will be valid."
'  fpcmbYear3.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbYear4.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbPenMonth.ToolTipText = "Select the last month which the new business license will be valid."
'  fpcmbPenDay.ToolTipText = "Select the last day which the new business license will be valid."
'  fpcmbYear5.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbWholeMonth.ToolTipText = "Select the last month to include for total receipts for last year for Wholesale Merchants."
'  fpcmbWholeDay.ToolTipText = "Select the last day to include for total receipts for last year for Wholesale Merchants."
'  fpcmbYear6.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbRetailMonth.ToolTipText = "Select the last month to include for total receipts for last year for Retail."
'  fpcmbRetailDay.ToolTipText = "Select the last day to include for total receipts for last year for Retail."
'  fpcmbYear7.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbFinMonth.ToolTipText = "Select the last month to include for total receipts for last year for Financial, Real Estate and Professional services."
'  fpcmbFinDay.ToolTipText = "Select the last day to include for total receipts for last year for Financial, Real Estate and Professional services."
'  fpcmbYear8.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbContMonth.ToolTipText = "Select the last month to include for total receipts for last year for Contracting."
'  fpcmbContDay.ToolTipText = "Select the last day to include for total receipts for last year for Contracting."
'  fpcmbYear9.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fpcmbRepairMonth.ToolTipText = "Select the last month to include for total receipts for last year for Repair, Personal or Business Services."
'  fpcmbRepairDay.ToolTipText = "Select the last day to include for total receipts for last year for Repair, Personal or Business Services."
'  fpcmbYear10.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtPhone.ToolTipText = "Enter the town's official phone number here."
'  fpcmbAppRetMonth.ToolTipText = "Select the month by which Business License renewal applications must be returned."
'  fpcmbAppRetDay.ToolTipText = "Select the day by which Business License renewal applications must be returned."
'  fpcmbLicRetMonth.ToolTipText = "Select the month by which Business License fees must be paid."
'  fpcmbLicRetDay.ToolTipText = "Select the day by which Business License fees must be paid."
'  cmdNext.ToolTipText = "Press to move to application template #5."
'  cmdLast.ToolTipText = "Press to move to business application #3."
'  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'  cmdSave.ToolTipText = "Press to save the data on this screen."
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec

    If QPTrim(TownRec.AppTownOf) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtTownName.Text) <> "" Then
        fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
      Else
        fptxtTownOf.Text = "Town Of 'Your Town'"
      End If
    Else
      fptxtTownOf.Text = QPTrim(TownRec.AppTownOf)
    End If

    If QPTrim(TownRec.AppAdd1) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) <> "" Then
        fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
      Else
        fptxtAdd.Text = "Street Address"
      End If
    Else
      fptxtAdd.Text = QPTrim(TownRec.AppAdd1)
    End If

    If QPTrim(TownRec.AppCity) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtCity.Text) <> "" Then
        fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
      Else
        fptxtCity.Text = "Your Town"
      End If
    Else
      fptxtCity.Text = QPTrim(TownRec.AppCity)
    End If

    If QPTrim(TownRec.AppState) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtState.Text) <> "" Then
        fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
      Else
        fptxtState.Text = "ST"
      End If
    Else
      fptxtState.Text = QPTrim(TownRec.AppState)
    End If

    If QPTrim(TownRec.AppZip) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtZip.Text) <> "" Then
        fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
      Else
        fptxtZip.Text = "11111-1111"
      End If
    Else
      fptxtZip.Text = QPTrim(TownRec.AppZip)
    End If

    If QPTrim(TownRec.AppPhone) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtPhone.Text) <> "(" Then
        fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
      Else
        fptxtPhone.Text = "(111)111-1111"
      End If
    Else
      fptxtPhone.Text = QPTrim(TownRec.AppPhone)
    End If
    
    If QPTrim$(TownRec.AppMayorCouncil) <> "" Then
      fptxtTownBody.Text = QPTrim$(TownRec.AppMayorCouncil)
    Else
      fptxtTownBody.Text = "Official Town Body"
    End If
    
    If TownRec.AppWholeMonth > 0 Then
      fpcmbWholeMonth.Text = TownRec.AppWholeMonth
    Else
      fpcmbWholeMonth.Text = "1"
    End If
    
    If TownRec.AppWholeDay > 0 Then
      fpcmbWholeDay.Text = TownRec.AppWholeDay
    Else
      fpcmbWholeDay.Text = "1"
    End If
    
    If TownRec.AppRetailMonth > 0 Then
      fpcmbRetailMonth.Text = TownRec.AppRetailMonth
    Else
      fpcmbRetailMonth.Text = "1"
    End If
    
    If TownRec.AppRetailDay > 0 Then
      fpcmbRetailDay.Text = TownRec.AppRetailDay
    Else
      fpcmbRetailDay.Text = "1"
    End If
    
    If TownRec.AppFinMonth > 0 Then
      fpcmbFinMonth.Text = TownRec.AppFinMonth
    Else
      fpcmbFinMonth.Text = "1"
    End If
    
    If TownRec.AppFinDay > 0 Then
      fpcmbFinDay.Text = TownRec.AppFinDay
    Else
      fpcmbFinDay.Text = "1"
    End If
    
    If TownRec.AppContMonth > 0 Then
      fpcmbContMonth.Text = TownRec.AppContMonth
    Else
      fpcmbContMonth.Text = "1"
    End If
    
    If TownRec.AppContDay > 0 Then
      fpcmbContDay.Text = TownRec.AppContDay
    Else
      fpcmbContDay.Text = "1"
    End If
    
    If TownRec.AppRepairMonth > 0 Then
      fpcmbRepairMonth.Text = TownRec.AppRepairMonth
    Else
      fpcmbRepairMonth.Text = "1"
    End If
    
    If TownRec.AppRepairDay > 0 Then
      fpcmbRepairDay.Text = TownRec.AppRepairDay
    Else
      fpcmbRepairDay.Text = "1"
    End If
    
    If QPTrim$(TownRec.AppPenMonth) <> "" Then
      fpcmbPenMonth.Text = QPTrim$(TownRec.AppPenMonth)
    Else
      fpcmbPenMonth.Text = "January"
    End If
    
    If TownRec.AppPenDay > 0 Then
      fpcmbPenDay.Text = TownRec.AppPenDay
    Else
      fpcmbPenDay.Text = "1"
    End If
    
    If QPTrim$(TownRec.AppFiscMonth) <> "" Then
      fpcmbAppRetMonth.Text = UCase(QPTrim$(TownRec.AppFiscMonth))
    Else
      fpcmbAppRetMonth.Text = "JANUARY"
    End If
    
    If TownRec.AppFiscDay > 0 Then
      fpcmbAppRetDay.Text = TownRec.AppFiscDay
    Else
      fpcmbAppRetDay.Text = "1"
    End If
    
    If QPTrim$(TownRec.AppStartMonth) <> "" Then
      fpcmbStartMonth.Text = QPTrim$(TownRec.AppStartMonth)
    Else
      fpcmbStartMonth.Text = "January"
    End If
    
    If TownRec.AppStartDay > 0 Then
      fpcmbStartDay.Text = TownRec.AppStartDay
    Else
      fpcmbStartDay.Text = "1"
    End If
    
    If QPTrim$(TownRec.AppLicRetMonth) <> "" Then
      fpcmbLicRetMonth.Text = UCase(QPTrim$(TownRec.AppLicRetMonth))
    Else
      fpcmbLicRetMonth.Text = "JANUARY"
    End If
    
    If TownRec.AppLicRetDay > 0 Then
      fpcmbLicRetDay.Text = TownRec.AppLicRetDay
    Else
      fpcmbLicRetDay.Text = "1"
    End If
    
    If TownRec.AppAdoptDate > 0 Then
      fptxtAdoptDate.Text = MakeRegDate(TownRec.AppAdoptDate)
    Else
      fptxtAdoptDate.Text = Date$
    End If
    
    If QPTrim$(TownRec.AppDiscMonth) <> "" Then
      fpcmbDiscMonth.Text = QPTrim$(TownRec.AppDiscMonth)
    Else
      fpcmbDiscMonth.Text = "JAN"
    End If
    
    If TownRec.AppDiscDay > 0 Then
      fpcmbDiscDay.Text = TownRec.AppDiscDay
    Else
      fpcmbDiscDay.Text = "1"
    End If
    
    If QPTrim$(TownRec.AppCityOrd) <> "" Then
      fptxtOrdinance.Text = QPTrim$(TownRec.AppCityOrd)
    Else
      fptxtOrdinance.Text = ""
    End If
    
    For x = 1 To 10
      If QPTrim$(TownRec.AppYrUpDown(x)) = "0" Then TownRec.AppYrUpDown(x) = "Curr"
    Next x
    
    fpcmbYear1.Text = QPTrim$(TownRec.AppYrUpDown(1))
    fpcmbYear2.Text = QPTrim$(TownRec.AppYrUpDown(2))
    fpcmbYear3.Text = QPTrim$(TownRec.AppYrUpDown(3))
    fpcmbYear4.Text = QPTrim$(TownRec.AppYrUpDown(4))
    fpcmbYear5.Text = QPTrim$(TownRec.AppYrUpDown(5))
    fpcmbYear6.Text = QPTrim$(TownRec.AppYrUpDown(6))
    fpcmbYear7.Text = QPTrim$(TownRec.AppYrUpDown(7))
    fpcmbYear8.Text = QPTrim$(TownRec.AppYrUpDown(8))
    fpcmbYear9.Text = QPTrim$(TownRec.AppYrUpDown(9))
    fpcmbYear10.Text = QPTrim$(TownRec.AppYrUpDown(10))
    
    lblPage2TownHead.Caption = QPTrim$(fptxtTownOf.Text)
    lblPenTownname.Caption = QPTrim$(fptxtTownOf.Text)
    lblRespectTown.Caption = QPTrim$(fptxtTownOf.Text)
    lblOrdCity.Caption = QPTrim$(fptxtCity.Text)
    lblTownOrd.Caption = UCase(QPTrim$(fptxtOrdinance.Text))
    lblTownName2.Caption = QPTrim$(fptxtTownOf.Text)
  Else
    If QPTrim$(frmBLTownSetup.fptxtTownName.Text) <> "" Then
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    Else
      fptxtTownOf.Text = "Town Of 'Your Town'"
    End If

    lblPage2TownHead.Caption = QPTrim$(fptxtTownOf.Text)
    lblPenTownname.Caption = QPTrim$(fptxtTownOf.Text)
    lblRespectTown.Caption = QPTrim$(fptxtTownOf.Text)
    If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) <> "" Then
      fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    Else
      fptxtAdd.Text = "Street Address"
    End If

    If QPTrim$(frmBLTownSetup.fptxtCity.Text) <> "" Then
      fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    Else
      fptxtCity.Text = "Your Town"
    End If
    
    lblOrdCity.Caption = QPTrim$(fptxtCity.Text)
    
    If QPTrim$(frmBLTownSetup.fptxtState.Text) <> "" Then
      fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
    Else
      fptxtState.Text = "ST"
    End If

    If QPTrim$(frmBLTownSetup.fptxtZip.Text) <> "" Then
      fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
    Else
      fptxtZip.Text = "11111-1111"
    End If

    If QPTrim$(frmBLTownSetup.fptxtPhone.Text) <> "(" Then
      fptxtPhone.Text = QPTrim$(frmBLTownSetup.fptxtPhone.Text)
    Else
      fptxtPhone.Text = "(111)111-1111"
    End If
    
    fptxtTownBody.Text = "Official Town Body"
    fpcmbWholeMonth.Text = "1"
    fpcmbWholeDay.Text = "1"
    fpcmbRetailMonth.Text = "1"
    fpcmbRetailDay.Text = "1"
    fpcmbFinMonth.Text = "1"
    fpcmbFinDay.Text = "1"
    fpcmbContMonth.Text = "1"
    fpcmbContDay.Text = "1"
    fpcmbRepairMonth.Text = "1"
    fpcmbRepairDay.Text = "1"
    fpcmbPenMonth.Text = "January"
    fpcmbPenDay.Text = "1"
    fpcmbAppRetMonth.Text = "JANUARY"
    fpcmbAppRetDay.Text = "1"
    fpcmbStartMonth.Text = "January"
    fpcmbStartDay.Text = "1"
    fpcmbLicRetMonth.Text = "JANUARY"
    fpcmbLicRetDay.Text = "1"
    fptxtAdoptDate.Text = Date$
    fpcmbDiscMonth.Text = "JAN"
    fpcmbDiscDay.Text = "1"
    fptxtOrdinance.Text = "TOWN ORDINANACE #XXX"
    lblPage2TownHead.Caption = QPTrim$(fptxtTownOf.Text)
    lblPenTownname.Caption = QPTrim$(fptxtTownOf.Text)
    lblRespectTown.Caption = QPTrim$(fptxtTownOf.Text)
    lblOrdCity.Caption = UCase(QPTrim$(fptxtCity.Text))
    lblTownOrd.Caption = QPTrim$(fptxtOrdinance.Text)
    lblTownName2.Caption = QPTrim$(fptxtTownOf.Text)
    fpcmbYear1.Text = "Curr"
    fpcmbYear2.Text = "Curr"
    fpcmbYear3.Text = "Curr"
    fpcmbYear4.Text = "Curr"
    fpcmbYear5.Text = "Curr"
    fpcmbYear6.Text = "Curr"
    fpcmbYear7.Text = "Curr"
    fpcmbYear8.Text = "Curr"
    fpcmbYear9.Text = "Curr"
    fpcmbYear10.Text = "Curr"
    
  End If
  Close THandle

  For x = 1 To 12
    Select Case x
      Case 1
        fpcmbLicRetMonth.AddItem "JANUARY"
        fpcmbStartMonth.AddItem "January"
        fpcmbAppRetMonth.AddItem "JANUARY"
        fpcmbPenMonth.AddItem "January"
        fpcmbDiscMonth.AddItem "JAN"
        fpcmbWholeMonth.AddItem "1"
        fpcmbRetailMonth.AddItem "1"
        fpcmbFinMonth.AddItem "1"
        fpcmbContMonth.AddItem "1"
        fpcmbRepairMonth.AddItem "1"
      Case 2
        fpcmbLicRetMonth.AddItem "FEBRUARY"
        fpcmbStartMonth.AddItem "February"
        fpcmbAppRetMonth.AddItem "FEBRUARY"
        fpcmbPenMonth.AddItem "February"
        fpcmbDiscMonth.AddItem "FEB"
        fpcmbWholeMonth.AddItem "2"
        fpcmbRetailMonth.AddItem "2"
        fpcmbFinMonth.AddItem "2"
        fpcmbContMonth.AddItem "2"
        fpcmbRepairMonth.AddItem "2"
      Case 3
        fpcmbLicRetMonth.AddItem "MARCH"
        fpcmbStartMonth.AddItem "March"
        fpcmbAppRetMonth.AddItem "MARCH"
        fpcmbPenMonth.AddItem "March"
        fpcmbDiscMonth.AddItem "MAR"
        fpcmbWholeMonth.AddItem "3"
        fpcmbRetailMonth.AddItem "3"
        fpcmbFinMonth.AddItem "3"
        fpcmbContMonth.AddItem "3"
        fpcmbRepairMonth.AddItem "3"
      Case 4
        fpcmbLicRetMonth.AddItem "APRIL"
        fpcmbStartMonth.AddItem "April"
        fpcmbAppRetMonth.AddItem "APRIL"
        fpcmbPenMonth.AddItem "April"
        fpcmbDiscMonth.AddItem "APR"
        fpcmbWholeMonth.AddItem "4"
        fpcmbRetailMonth.AddItem "4"
        fpcmbFinMonth.AddItem "4"
        fpcmbContMonth.AddItem "4"
        fpcmbRepairMonth.AddItem "4"
      Case 5
        fpcmbLicRetMonth.AddItem "MAY"
        fpcmbStartMonth.AddItem "May"
        fpcmbAppRetMonth.AddItem "MAY"
        fpcmbPenMonth.AddItem "May"
        fpcmbDiscMonth.AddItem "MAY"
        fpcmbWholeMonth.AddItem "5"
        fpcmbRetailMonth.AddItem "5"
        fpcmbFinMonth.AddItem "5"
        fpcmbContMonth.AddItem "5"
        fpcmbRepairMonth.AddItem "5"
      Case 6
        fpcmbLicRetMonth.AddItem "JUNE"
        fpcmbStartMonth.AddItem "June"
        fpcmbAppRetMonth.AddItem "JUNE"
        fpcmbPenMonth.AddItem "June"
        fpcmbDiscMonth.AddItem "JUN"
        fpcmbWholeMonth.AddItem "6"
        fpcmbRetailMonth.AddItem "6"
        fpcmbFinMonth.AddItem "6"
        fpcmbContMonth.AddItem "6"
        fpcmbRepairMonth.AddItem "6"
      Case 7
        fpcmbLicRetMonth.AddItem "JULY"
        fpcmbStartMonth.AddItem "July"
        fpcmbAppRetMonth.AddItem "JULY"
        fpcmbPenMonth.AddItem "July"
        fpcmbDiscMonth.AddItem "JUL"
        fpcmbWholeMonth.AddItem "7"
        fpcmbRetailMonth.AddItem "7"
        fpcmbFinMonth.AddItem "7"
        fpcmbContMonth.AddItem "7"
        fpcmbRepairMonth.AddItem "7"
      Case 8
        fpcmbLicRetMonth.AddItem "AUGUST"
        fpcmbStartMonth.AddItem "August"
        fpcmbAppRetMonth.AddItem "AUGUST"
        fpcmbPenMonth.AddItem "August"
        fpcmbDiscMonth.AddItem "AUG"
        fpcmbWholeMonth.AddItem "8"
        fpcmbRetailMonth.AddItem "8"
        fpcmbFinMonth.AddItem "8"
        fpcmbContMonth.AddItem "8"
        fpcmbRepairMonth.AddItem "8"
      Case 9
        fpcmbLicRetMonth.AddItem "SEPTEMBER"
        fpcmbStartMonth.AddItem "September"
        fpcmbAppRetMonth.AddItem "SEPTEMBER"
        fpcmbPenMonth.AddItem "September"
        fpcmbDiscMonth.AddItem "SEP"
        fpcmbWholeMonth.AddItem "9"
        fpcmbRetailMonth.AddItem "9"
        fpcmbFinMonth.AddItem "9"
        fpcmbContMonth.AddItem "9"
        fpcmbRepairMonth.AddItem "9"
      Case 10
        fpcmbLicRetMonth.AddItem "OCTOBER"
        fpcmbStartMonth.AddItem "October"
        fpcmbAppRetMonth.AddItem "OCTOBER"
        fpcmbPenMonth.AddItem "October"
        fpcmbDiscMonth.AddItem "OCT"
        fpcmbWholeMonth.AddItem "10"
        fpcmbRetailMonth.AddItem "10"
        fpcmbFinMonth.AddItem "10"
        fpcmbContMonth.AddItem "10"
        fpcmbRepairMonth.AddItem "10"
      Case 11
        fpcmbLicRetMonth.AddItem "NOVEMBER"
        fpcmbStartMonth.AddItem "November"
        fpcmbAppRetMonth.AddItem "NOVEMBER"
        fpcmbPenMonth.AddItem "November"
        fpcmbDiscMonth.AddItem "NOV"
        fpcmbWholeMonth.AddItem "11"
        fpcmbRetailMonth.AddItem "11"
        fpcmbFinMonth.AddItem "11"
        fpcmbContMonth.AddItem "11"
        fpcmbRepairMonth.AddItem "11"
      Case 12
        fpcmbLicRetMonth.AddItem "DECEMBER"
        fpcmbStartMonth.AddItem "December"
        fpcmbAppRetMonth.AddItem "DECEMBER"
        fpcmbPenMonth.AddItem "December"
        fpcmbDiscMonth.AddItem "DEC"
        fpcmbWholeMonth.AddItem "12"
        fpcmbRetailMonth.AddItem "12"
        fpcmbFinMonth.AddItem "12"
        fpcmbContMonth.AddItem "12"
        fpcmbRepairMonth.AddItem "12"
    End Select
  Next x

  For x = 1 To 31
    fpcmbLicRetDay.AddItem CStr(x)
    fpcmbAppRetDay.AddItem CStr(x)
    fpcmbPenDay.AddItem CStr(x)
    fpcmbContDay.AddItem CStr(x)
    fpcmbRepairDay.AddItem CStr(x)
    fpcmbRetailDay.AddItem CStr(x)
    fpcmbStartDay.AddItem CStr(x)
    fpcmbWholeDay.AddItem CStr(x)
    fpcmbFinDay.AddItem CStr(x)
    fpcmbDiscDay.AddItem CStr(x)
  Next x
  
  fpcmbYear1.AddItem "Curr"
  fpcmbYear1.AddItem "+1"
  fpcmbYear1.AddItem "-1"
  fpcmbYear2.AddItem "Curr"
  fpcmbYear2.AddItem "+1"
  fpcmbYear2.AddItem "-1"
  fpcmbYear3.AddItem "Curr"
  fpcmbYear3.AddItem "+1"
  fpcmbYear3.AddItem "-1"
  fpcmbYear4.AddItem "Curr"
  fpcmbYear4.AddItem "+1"
  fpcmbYear4.AddItem "-1"
  fpcmbYear5.AddItem "Curr"
  fpcmbYear5.AddItem "+1"
  fpcmbYear5.AddItem "-1"
  fpcmbYear6.AddItem "Curr"
  fpcmbYear6.AddItem "+1"
  fpcmbYear6.AddItem "-1"
  fpcmbYear7.AddItem "Curr"
  fpcmbYear7.AddItem "+1"
  fpcmbYear7.AddItem "-1"
  fpcmbYear8.AddItem "Curr"
  fpcmbYear8.AddItem "+1"
  fpcmbYear8.AddItem "-1"
  fpcmbYear9.AddItem "Curr"
  fpcmbYear9.AddItem "+1"
  fpcmbYear9.AddItem "-1"
  fpcmbYear10.AddItem "Curr"
  fpcmbYear10.AddItem "+1"
  fpcmbYear10.AddItem "-1"
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate4", "LoadMe", Erl)
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

Private Sub fpcmbAppRetDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 1
  fpcmbAppRetDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbAppRetDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAppRetDay.ListIndex = -1
  End If
  If fpcmbAppRetDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLicRetMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbAppRetMonth_Change()
  fpcmbAppRetMonth.Text = UCase(fpcmbAppRetMonth.Text)
End Sub

Private Sub fpcmbAppRetMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 1
  fpcmbAppRetMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbAppRetMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAppRetMonth.ListIndex = -1
  End If
  If fpcmbAppRetMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbAppRetDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbContDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbContDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbContDay.ListIndex = -1
  End If
  If fpcmbContDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear9.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbContMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbContMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbContMonth.ListIndex = -1
  End If
  If fpcmbContMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbContDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbDiscDay_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fpcmbDiscDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbDiscDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscDay.ListIndex = -1
  End If
  If fpcmbDiscDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbDiscMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fpcmbDiscMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbDiscMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscMonth.ListIndex = -1
  End If
  If fpcmbDiscMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDiscDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLicRetDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 1
  fpcmbLicRetDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbLicRetDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLicRetDay.ListIndex = -1
  End If
  If fpcmbLicRetDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      vaTabPro1.ActiveTab = 0
      fptxtTownOf.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLicRetMonth_Change()
  fpcmbLicRetMonth.Text = UCase(fpcmbLicRetMonth.Text)

End Sub

Private Sub fpcmbRepairMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbRepairMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRepairMonth.ListIndex = -1
  End If
  If fpcmbRepairMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRepairDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRepairDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbRepairDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRepairDay.ListIndex = -1
  End If
  If fpcmbRepairDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear10.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLicRetMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 1
  fpcmbLicRetMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbLicRetMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLicRetMonth.ListIndex = -1
  End If
  If fpcmbLicRetMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLicRetDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPenMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 0
  fpcmbPenMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbPenMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPenMonth.ListIndex = -1
  End If
  If fpcmbPenMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPenDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRetailMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbRetailMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRetailMonth.ListIndex = -1
  End If
  If fpcmbRetailMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRepairDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRetailDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbRetailDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRetailDay.ListIndex = -1
  End If
  If fpcmbRetailDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear7.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStartMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 0
  fpcmbStartMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbStartMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStartMonth.ListIndex = -1
  End If
  If fpcmbStartMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbStartDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStartDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 0
  fpcmbStartDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbStartDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStartDay.ListIndex = -1
  End If
  If fpcmbStartDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear3.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbWholeMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbWholeMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbWholeMonth.ListIndex = -1
  End If
  If fpcmbWholeMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbWholeDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbWholeDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbWholeDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbWholeDay.ListIndex = -1
  End If
  If fpcmbWholeDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear6.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbFinDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbFinDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFinDay.ListIndex = -1
  End If
  If fpcmbFinDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear8.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbFinMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbFinMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFinMonth.ListIndex = -1
  End If
  If fpcmbFinMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbFinDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPenDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  vaTabPro1.ActiveTab = 0
  fpcmbPenDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbPenDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPenDay.ListIndex = -1
  End If
  If fpcmbPenDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear5.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear1_Change()
  If QPTrim$(fpcmbYear1.Text) = "Curr" Then
    Label80.Caption = "For Year: Curr"
  ElseIf QPTrim$(fpcmbYear1.Text) = "+1" Then
    Label80.Caption = "For Year: +1"
  ElseIf QPTrim$(fpcmbYear1.Text) = "-1" Then
    Label80.Caption = "For Year: -1"
  Else
    Label80.Caption = "For Year: 0000"
  End If
End Sub

Private Sub fptxtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtAdd.BackColor = -2147483643

End Sub

Private Sub fptxtAdoptDate_Change()
  vaTabPro1.ActiveTab = 0
  fptxtAdoptDate.BackColor = -2147483643

End Sub

Private Sub fptxtCity_Change()
  lblOrdCity.Caption = fptxtCity.Text

End Sub

Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtCity.BackColor = -2147483643

End Sub

Private Sub fptxtOrdinance_Change()
  lblTownOrd.Caption = fptxtOrdinance.Text

End Sub

Private Sub fptxtOrdinance_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtOrdinance.BackColor = -2147483643

End Sub

Private Sub fptxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtPhone.BackColor = -2147483643

End Sub

Private Sub fptxtState_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtState.BackColor = -2147483643

End Sub

Private Sub fptxtTownBody_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtTownBody.BackColor = -2147483643

End Sub

Private Sub fptxtTownOf_Change()
  lblPage2TownHead.Caption = QPTrim$(fptxtTownOf.Text)
  lblPenTownname.Caption = QPTrim$(fptxtTownOf.Text)
  lblRespectTown.Caption = QPTrim$(fptxtTownOf.Text)
  lblTownName2.Caption = QPTrim$(fptxtTownOf.Text)
End Sub

Private Sub fptxtTownOf_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtTownOf.BackColor = -2147483643

End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  vaTabPro1.ActiveTab = 0
  fptxtZip.BackColor = -2147483643

End Sub

Private Sub cmdNext_Click()
  frmBLAppTemplate5.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLast_Click()
  frmBLAppTemplate3.Show
  DoEvents
  Unload Me
End Sub

Private Sub fpcmbYear1_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear1.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear1.ListIndex = -1
  End If
  If fpcmbYear1.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtOrdinance.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear2_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear2.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear2.ListIndex = -1
  End If
  If fpcmbYear2.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtTownBody.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear3_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear3.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear3.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear3.ListIndex = -1
  End If
  If fpcmbYear3.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear4.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear4_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear4.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear4.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear4.ListIndex = -1
  End If
  If fpcmbYear4.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPenMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear5_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear5.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear5.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear5.ListIndex = -1
  End If
  If fpcmbYear5.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      vaTabPro1.ActiveTab = 1
      fpcmbWholeMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear6_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear6.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear6.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear6.ListIndex = -1
  End If
  If fpcmbYear6.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRetailMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear7_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear7.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear7.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear7.ListIndex = -1
  End If
  If fpcmbYear7.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbFinMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear8_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear8.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear8.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear8.ListIndex = -1
  End If
  If fpcmbYear8.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbContMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear9_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear9.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear9.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear9.ListIndex = -1
  End If
  If fpcmbYear9.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRepairMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear10_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear10.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear10.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear10.ListIndex = -1
  End If
  If fpcmbYear10.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPhone.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

