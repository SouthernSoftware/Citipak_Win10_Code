VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Begin VB.Form frmW2ElecSub 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License W2 Electronic Submission"
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
   Icon            =   "frmW2ElecSub.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6684
      Left            =   492
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1104
      Width           =   10668
      _Version        =   196609
      _ExtentX        =   18817
      _ExtentY        =   11790
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
      TabsPerRow      =   2
      TabCount        =   2
      Tab             =   1
      AlignTextH      =   1
      GrayAreaColor   =   13684944
      OffsetFromClientTop=   -1  'True
      DataFormat      =   ""
      AutoSizeChildren=   3
      BookCornerGuardWidth=   90
      BookCornerGuardLength=   375
      DataField       =   ""
      TabCaption      =   "frmW2ElecSub.frx":08CA
      PageEarMarkPictureNext=   "frmW2ElecSub.frx":0B59
      PageEarMarkPicturePrev=   "frmW2ElecSub.frx":0B75
      EarMarkPictureNext=   "frmW2ElecSub.frx":0B91
      EarMarkPicturePrev=   "frmW2ElecSub.frx":0BAD
      Begin VB.Frame Frame2 
         BackColor       =   &H008F8265&
         BorderStyle     =   0  'None
         Height          =   5715
         Left            =   360
         TabIndex        =   64
         Top             =   585
         Width           =   10020
         Begin LpLib.fpCombo fpcmbAgentCode 
            Height          =   375
            Left            =   3030
            TabIndex        =   29
            ToolTipText     =   $"frmW2ElecSub.frx":0BC9
            Top             =   1155
            Width           =   2985
            _Version        =   196608
            _ExtentX        =   5265
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
            ColDesigner     =   "frmW2ElecSub.frx":0C68
         End
         Begin LpLib.fpCombo fpcmbTermBusInd 
            Height          =   375
            Left            =   6810
            TabIndex        =   33
            ToolTipText     =   "Enter '1' if the town was terminated as a business entity during this tax year. Otherwise enter '0'."
            Top             =   2205
            Width           =   2205
            _Version        =   196608
            _ExtentX        =   3889
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
            ColDesigner     =   "frmW2ElecSub.frx":0F97
         End
         Begin LpLib.fpCombo fpcmbEmprState 
            Height          =   375
            Left            =   5805
            TabIndex        =   38
            ToolTipText     =   "Select the employer's state."
            Top             =   4170
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ColDesigner     =   "frmW2ElecSub.frx":12C6
         End
         Begin LpLib.fpCombo fpcmb3rdSckPay 
            Height          =   375
            Left            =   5520
            TabIndex        =   41
            ToolTipText     =   "Enter '1' for a sick pay indicator. Otherwise, enter '0'."
            Top             =   5085
            Width           =   1980
            _Version        =   196608
            _ExtentX        =   3492
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
            ColDesigner     =   "frmW2ElecSub.frx":15F5
         End
         Begin LpLib.fpCombo fpcmbEmpKind 
            Height          =   375
            Left            =   2505
            TabIndex        =   26
            Top             =   240
            Width           =   2985
            _Version        =   196608
            _ExtentX        =   5265
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
            ColDesigner     =   "frmW2ElecSub.frx":1924
         End
         Begin EditLib.fpText fptxtEmprAgtEIN 
            Height          =   348
            Left            =   6624
            TabIndex        =   31
            ToolTipText     =   $"frmW2ElecSub.frx":1C53
            Top             =   1680
            Width           =   1356
            _Version        =   196608
            _ExtentX        =   2392
            _ExtentY        =   614
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
            MaxLength       =   9
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
         Begin EditLib.fpDateTime fptxtTaxYear 
            Height          =   375
            Left            =   7680
            TabIndex        =   28
            ToolTipText     =   "Enter the Tax Year for which this file will be submitted."
            Top             =   240
            Width           =   1110
            _Version        =   196608
            _ExtentX        =   1968
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
            AlignTextH      =   1
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
            Text            =   "2019"
            DateCalcMethod  =   1
            DateTimeFormat  =   5
            UserDefinedFormat=   "yyyy"
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
            ButtonColor     =   13684944
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fptxtAgent4EIN 
            Height          =   348
            Left            =   7632
            TabIndex        =   30
            ToolTipText     =   $"frmW2ElecSub.frx":1CE1
            Top             =   1152
            Width           =   1356
            _Version        =   196608
            _ExtentX        =   2392
            _ExtentY        =   614
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
            MaxLength       =   9
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
         Begin EditLib.fpText fptxtOtherEIN 
            Height          =   348
            Left            =   2112
            TabIndex        =   32
            ToolTipText     =   $"frmW2ElecSub.frx":1D6C
            Top             =   2256
            Width           =   1356
            _Version        =   196608
            _ExtentX        =   2392
            _ExtentY        =   614
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
            MaxLength       =   9
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
         Begin EditLib.fpText fptxtEmployerName 
            Height          =   348
            Left            =   1968
            TabIndex        =   34
            ToolTipText     =   "Enter the name associated with the EIN entered in the Employer/Agent Employer ID# field."
            Top             =   3120
            Width           =   7740
            _Version        =   196608
            _ExtentX        =   13652
            _ExtentY        =   614
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
            MaxLength       =   57
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
         Begin EditLib.fpText fptxtEmprAdd1 
            Height          =   348
            Left            =   1968
            TabIndex        =   35
            ToolTipText     =   "Enter the employer's delivery address (Street or Post Office Box)."
            Top             =   3648
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtEmprAdd2 
            Height          =   348
            Left            =   6744
            TabIndex        =   36
            ToolTipText     =   "Enter the employer's location address (Attention, Suite, Room Number, etc.)"
            Top             =   3648
            Width           =   2964
            _Version        =   196608
            _ExtentX        =   5228
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtEmprCity 
            Height          =   348
            Left            =   1968
            TabIndex        =   37
            ToolTipText     =   "Enter the town name."
            Top             =   4176
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtEmprZip 
            Height          =   348
            Left            =   7680
            TabIndex        =   39
            ToolTipText     =   "Enter the employer's state."
            Top             =   4176
            Width           =   1164
            _Version        =   196608
            _ExtentX        =   2053
            _ExtentY        =   614
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
            MaxLength       =   5
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
         Begin EditLib.fpText fptxtEmprZipX 
            Height          =   348
            Left            =   8976
            TabIndex        =   40
            ToolTipText     =   "If available, enter the zip extension number."
            Top             =   4176
            Width           =   732
            _Version        =   196608
            _ExtentX        =   1291
            _ExtentY        =   614
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
            MaxLength       =   4
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
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Kind of Employer"
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
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   330
            Width           =   2100
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0080FFFF&
            X1              =   0
            X2              =   10032
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0080FFFF&
            X1              =   0
            X2              =   10032
            Y1              =   912
            Y2              =   912
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Third-Party Sick Pay Indicator"
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
            Height          =   252
            Left            =   2496
            TabIndex        =   77
            Top             =   5184
            Width           =   2892
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
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
            Left            =   7296
            TabIndex        =   76
            Top             =   4224
            Width           =   348
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   5232
            TabIndex        =   75
            Top             =   4272
            Width           =   540
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Left            =   1488
            TabIndex        =   74
            Top             =   4272
            Width           =   396
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Location Address"
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
            Left            =   5088
            TabIndex        =   73
            Top             =   3744
            Width           =   1644
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Address"
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
            Left            =   336
            TabIndex        =   72
            Top             =   3744
            Width           =   1596
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Employer Name"
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
            Height          =   252
            Left            =   312
            TabIndex        =   71
            Top             =   3216
            Width           =   1572
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Other EIN"
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
            Height          =   252
            Left            =   1008
            TabIndex        =   70
            Top             =   2352
            Width           =   1020
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Terminating Business Indicator"
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
            Height          =   252
            Left            =   3696
            TabIndex        =   69
            Top             =   2304
            Width           =   3084
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent for EIN"
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
            Height          =   252
            Left            =   6240
            TabIndex        =   68
            Top             =   1248
            Width           =   1308
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Employer/Agent Employer Identification Number (EIN)"
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
            Height          =   252
            Left            =   1272
            TabIndex        =   67
            Top             =   1776
            Width           =   5268
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Indicator Code"
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
            Height          =   252
            Left            =   888
            TabIndex        =   66
            Top             =   1248
            Width           =   2100
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Year"
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
            Height          =   255
            Left            =   6675
            TabIndex        =   65
            Top             =   330
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H008F8265&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5715
         Left            =   -25065
         TabIndex        =   42
         Top             =   -21165
         Width           =   10020
         Begin LpLib.fpCombo fpcmbSubState 
            Height          =   375
            Left            =   6090
            TabIndex        =   17
            ToolTipText     =   "Enter the submitter's state."
            Top             =   3450
            Width           =   600
            _Version        =   196608
            _ExtentX        =   1058
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
            ColDesigner     =   "frmW2ElecSub.frx":1E1C
         End
         Begin LpLib.fpCombo fpcmbState 
            Height          =   390
            Left            =   6090
            TabIndex        =   8
            ToolTipText     =   "Select the town's state."
            Top             =   2010
            Width           =   600
            _Version        =   196608
            _ExtentX        =   1058
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
            EditAlignH      =   0
            EditAlignV      =   0
            ColDesigner     =   "frmW2ElecSub.frx":214B
         End
         Begin LpLib.fpCombo fpcmbPrefMeth 
            Height          =   375
            Left            =   7995
            TabIndex        =   25
            ToolTipText     =   "Select one of the codes in the drop down box."
            Top             =   4890
            Width           =   1755
            _Version        =   196608
            _ExtentX        =   3096
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
            ColDesigner     =   "frmW2ElecSub.frx":247A
         End
         Begin LpLib.fpCombo fpcmbPrepCode 
            Height          =   375
            Left            =   4845
            TabIndex        =   27
            ToolTipText     =   "Select  who prepared this file. NOTE: IF MORE THAN ONE CODE APPLIES, USE THE ONE THAT BEST DESCRIBES WHO PREPARED THIS FILE."
            Top             =   5280
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
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
            ColDesigner     =   "frmW2ElecSub.frx":27A9
         End
         Begin LpLib.fpCombo fpcmbResub 
            Height          =   375
            Left            =   8115
            TabIndex        =   1
            ToolTipText     =   "Enter '1' if this file is being resubmitted. Otherwise select '0'."
            Top             =   195
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ColDesigner     =   "frmW2ElecSub.frx":2AD8
         End
         Begin EditLib.fpText fptxtSubAdd2 
            Height          =   348
            Left            =   6840
            TabIndex        =   15
            ToolTipText     =   "Enter the submitter's location address (Attention, Suite, Room Number, etc.)"
            Top             =   3072
            Width           =   2916
            _Version        =   196608
            _ExtentX        =   5143
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtEmpEIN 
            Height          =   348
            Left            =   4896
            TabIndex        =   0
            ToolTipText     =   "Enter the submitter's EIN. This EIN should match the EIN on the external label. Enter numbers only."
            Top             =   192
            Width           =   1404
            _Version        =   196608
            _ExtentX        =   2476
            _ExtentY        =   614
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
            MaxLength       =   9
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
         Begin EditLib.fpText fptxtPinNum 
            Height          =   348
            Left            =   4896
            TabIndex        =   2
            ToolTipText     =   "Enter the PIN assigned to the employee who is attesting to the accuracy of this file. Enter numbers only."
            Top             =   576
            Width           =   2364
            _Version        =   196608
            _ExtentX        =   4170
            _ExtentY        =   614
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
            MaxLength       =   8
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
         Begin EditLib.fpText fptxtTownName 
            Height          =   348
            Left            =   2016
            TabIndex        =   4
            ToolTipText     =   "Enter the name of the town to receive MMREF-1 annual filing instructions."
            Top             =   1248
            Width           =   7740
            _Version        =   196608
            _ExtentX        =   13652
            _ExtentY        =   614
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
            MaxLength       =   57
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
            Height          =   348
            Left            =   6840
            TabIndex        =   6
            ToolTipText     =   "Enter the town's location address (Attention, Suite, Room Number, etc.)"
            Top             =   1632
            Width           =   2916
            _Version        =   196608
            _ExtentX        =   5143
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtAdd1 
            Height          =   348
            Left            =   2016
            TabIndex        =   5
            ToolTipText     =   "Enter the town's delivery address (Street or Post Office Box)."
            Top             =   1632
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtCity 
            Height          =   348
            Left            =   2016
            TabIndex        =   7
            ToolTipText     =   "Enter the town name."
            Top             =   2016
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtZip 
            Height          =   348
            Left            =   7728
            TabIndex        =   9
            ToolTipText     =   "Enter the town's state."
            Top             =   2016
            Width           =   1164
            _Version        =   196608
            _ExtentX        =   2053
            _ExtentY        =   614
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   5
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
         Begin EditLib.fpText fptxtZipX 
            Height          =   348
            Left            =   9024
            TabIndex        =   10
            ToolTipText     =   "If available, enter the zip extension number."
            Top             =   2016
            Width           =   732
            _Version        =   196608
            _ExtentX        =   1291
            _ExtentY        =   614
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   4
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
         Begin EditLib.fpText fptxtSubmitterName 
            Height          =   348
            Left            =   2016
            TabIndex        =   11
            ToolTipText     =   "Enter the name of the organization to receive notification of unprocessable data."
            Top             =   2688
            Width           =   7740
            _Version        =   196608
            _ExtentX        =   13652
            _ExtentY        =   614
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
            MaxLength       =   57
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
         Begin EditLib.fpText fptxtSubAdd1 
            Height          =   348
            Left            =   2016
            TabIndex        =   13
            ToolTipText     =   "Enter the submitter's delivery address (Street or Post Office Box)."
            Top             =   3072
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtSubCity 
            Height          =   348
            Left            =   2016
            TabIndex        =   16
            ToolTipText     =   "Enter the submitter's town name."
            Top             =   3456
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtSubZip 
            Height          =   348
            Left            =   7728
            TabIndex        =   18
            ToolTipText     =   "Enter the submitter's state."
            Top             =   3456
            Width           =   1164
            _Version        =   196608
            _ExtentX        =   2053
            _ExtentY        =   614
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   5
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
         Begin EditLib.fpText fptxtSubZipX 
            Height          =   348
            Left            =   9024
            TabIndex        =   19
            ToolTipText     =   "If available, enter the zip extension number."
            Top             =   3456
            Width           =   732
            _Version        =   196608
            _ExtentX        =   1291
            _ExtentY        =   614
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
            CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
            MaxLength       =   4
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
         Begin EditLib.fpText fptxtCntctName 
            Height          =   348
            Left            =   2016
            TabIndex        =   20
            ToolTipText     =   "Enter the name of the person to be contacted by SSA concerning Processing problems."
            Top             =   4128
            Width           =   3036
            _Version        =   196608
            _ExtentX        =   5355
            _ExtentY        =   614
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
         Begin EditLib.fpMask fptxtPhone 
            Height          =   378
            Left            =   8088
            TabIndex        =   21
            ToolTipText     =   "Enter the contact's telephone number (including the area code)."
            Top             =   4108
            Width           =   1668
            _Version        =   196608
            _ExtentX        =   2942
            _ExtentY        =   667
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
         Begin EditLib.fpText fptxtPhoneX 
            Height          =   348
            Left            =   2856
            TabIndex        =   22
            ToolTipText     =   "Enter the contact's telephone extension."
            Top             =   4512
            Width           =   732
            _Version        =   196608
            _ExtentX        =   1291
            _ExtentY        =   614
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
            MaxLength       =   5
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
         Begin EditLib.fpText fptxtCntctEMail 
            Height          =   348
            Left            =   4368
            TabIndex        =   23
            ToolTipText     =   "If applicable, enter the contact's electronic mail/Internet address."
            Top             =   4512
            Width           =   5388
            _Version        =   196608
            _ExtentX        =   9504
            _ExtentY        =   614
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
            MaxLength       =   40
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
         Begin EditLib.fpMask fptxtCntctFax 
            Height          =   348
            Left            =   1632
            TabIndex        =   24
            ToolTipText     =   "If applicable, enter the contact's Fax number (including area code)."
            Top             =   4896
            Width           =   1716
            _Version        =   196608
            _ExtentX        =   3027
            _ExtentY        =   614
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
         Begin EditLib.fpText fptxtResubWFID 
            Height          =   348
            Left            =   8712
            TabIndex        =   3
            ToolTipText     =   "If you entered a '1' as a Resub Indicator then enter the WFID (Wage File Identifier) displayed on the notice sent to you by SSA."
            Top             =   576
            Width           =   1044
            _Version        =   196608
            _ExtentX        =   1841
            _ExtentY        =   614
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
            MaxLength       =   6
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
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "ReSub WFID"
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
            Left            =   7344
            TabIndex        =   79
            Top             =   672
            Width           =   1284
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "ReSub Indicator"
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
            Left            =   6456
            TabIndex        =   78
            Top             =   288
            Width           =   1572
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0080FFFF&
            X1              =   0
            X2              =   10032
            Y1              =   1056
            Y2              =   1056
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0080FFFF&
            X1              =   0
            X2              =   10032
            Y1              =   2496
            Y2              =   2496
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0080FFFF&
            X1              =   0
            X2              =   10032
            Y1              =   3984
            Y2              =   3984
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Preparer Code"
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
            Left            =   3384
            TabIndex        =   63
            Top             =   5376
            Width           =   1404
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Preferred Method of Problem Notification Code"
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
            Left            =   3480
            TabIndex        =   62
            Top             =   4992
            Width           =   4428
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact's Fax"
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
            Left            =   288
            TabIndex        =   61
            Top             =   4992
            Width           =   1356
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "EMail"
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
            Left            =   3696
            TabIndex        =   60
            Top             =   4608
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Submitter's Employer Identification Number (EIN)"
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
            Height          =   252
            Left            =   120
            TabIndex        =   59
            Top             =   288
            Width           =   4788
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Personal Identification Number (PIN)"
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
            Height          =   252
            Left            =   1272
            TabIndex        =   58
            Top             =   672
            Width           =   3540
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Town Name"
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
            Left            =   600
            TabIndex        =   57
            Top             =   1350
            Width           =   1236
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Location Address"
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
            Left            =   5136
            TabIndex        =   56
            Top             =   1728
            Width           =   1716
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Address"
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
            Left            =   288
            TabIndex        =   55
            Top             =   1776
            Width           =   1596
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Left            =   1440
            TabIndex        =   54
            Top             =   2112
            Width           =   540
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   5520
            TabIndex        =   53
            Top             =   2112
            Width           =   540
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
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
            Left            =   7296
            TabIndex        =   52
            Top             =   2064
            Width           =   348
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Submitter Name"
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
            Left            =   288
            TabIndex        =   51
            Top             =   2784
            Width           =   1644
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Address"
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
            TabIndex        =   50
            Top             =   3168
            Width           =   1596
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Left            =   1392
            TabIndex        =   49
            Top             =   3552
            Width           =   396
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   5520
            TabIndex        =   48
            Top             =   3504
            Width           =   540
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Zip"
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
            Left            =   7296
            TabIndex        =   47
            Top             =   3504
            Width           =   348
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Location Address"
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
            Left            =   5136
            TabIndex        =   46
            Top             =   3168
            Width           =   1716
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Name"
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
            Left            =   288
            TabIndex        =   45
            Top             =   4224
            Width           =   1596
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact's Telephone Number"
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
            Left            =   5256
            TabIndex        =   44
            Top             =   4224
            Width           =   2772
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact's Phone Extension"
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
            Left            =   288
            TabIndex        =   43
            Top             =   4608
            Width           =   2532
         End
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   5820
         Left            =   300
         Top             =   540
         Width           =   10125
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   5820
         Left            =   -25170
         Top             =   -21270
         Width           =   10125
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   5472
      TabIndex        =   80
      Top             =   8016
      Width           =   3210
      _Version        =   131072
      _ExtentX        =   5662
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
      ButtonDesigner  =   "frmW2ElecSub.frx":2E07
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   9108
      TabIndex        =   81
      Top             =   8016
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   868
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmW2ElecSub.frx":2FEF
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   588
      Index           =   1
      Left            =   1500
      Top             =   312
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "W-2 Electronic Submission"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2796
      TabIndex        =   14
      Top             =   456
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   708
      Left            =   1500
      Top             =   192
      Width           =   8652
   End
End
Attribute VB_Name = "frmW2ElecSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Dim DoWhatFlag As SaveChangeOptions1
  Dim changeFlag As Boolean
  Dim W2RARec As W2ElectronicSubRA
  Dim RAHandle As Integer
  Dim W2RERec As W2ElectronicSubRE
  Dim REHandle As Integer
  Dim PhoneNum$, FaxNum$
  
  If Exist("PRDATA\W2ESUBRA.DAT") Then
    OpenW2ESubRA RAHandle
    Get RAHandle, 1, W2RARec
    Close RAHandle
  Else
    GoTo NoRAFile
  End If
  
  If QPTrim$(fptxtEmpEIN.Text) <> QPTrim$(W2RARec.EINNum) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtEmpEIN.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtPinNum.Text) <> QPTrim$(W2RARec.PersIDNum) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtPinNum.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmbResub.Text, 1, 1) <> QPTrim$(W2RARec.ResubID) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbResub.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtResubWFID.Text) <> QPTrim$(W2RARec.ReSubWFID) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtResubWFID.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtTownName.Text) <> QPTrim$(W2RARec.CmpnyName) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtTownName.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtAdd1.Text) <> QPTrim$(W2RARec.DelAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtAdd1.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtAdd2.Text) <> QPTrim$(W2RARec.LocAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtAdd2.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtCity.Text) <> QPTrim$(W2RARec.City) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtCity.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fpcmbState.Text) <> QPTrim$(W2RARec.State) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbState.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtZip.Text) <> QPTrim$(W2RARec.Zip) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtZip.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtZipX.Text) <> QPTrim$(W2RARec.ZipExt) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtZipX.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubmitterName.Text) <> QPTrim$(W2RARec.SubmttrName) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubmitterName.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubAdd1.Text) <> QPTrim$(W2RARec.SubDelAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubAdd1.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubAdd2.Text) <> QPTrim$(W2RARec.SubLocAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubAdd2.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubCity.Text) <> QPTrim$(W2RARec.SubCity) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubCity.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fpcmbSubState.Text) <> QPTrim$(W2RARec.SubState) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbSubState.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubZip.Text) <> QPTrim$(W2RARec.SubZip) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubZip.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtSubZipX.Text) <> QPTrim$(W2RARec.SubZipExt) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtSubZipX.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtCntctName.Text) <> QPTrim$(W2RARec.ContactName) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtCntctName.SetFocus
    GoTo changeFound
  End If
  
  PhoneNum = ReplaceString(fptxtPhone.Text, "(", "")
  PhoneNum = ReplaceString(PhoneNum, ")", "")
  PhoneNum = ReplaceString(PhoneNum, "-", "")
  If QPTrim$(PhoneNum) <> QPTrim$(W2RARec.CntctPhone) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtPhone.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtPhoneX.Text) <> QPTrim$(W2RARec.CntPhnExt) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtPhoneX.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtCntctEMail.Text) <> QPTrim$(W2RARec.CntEMail) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtCntctEMail.SetFocus
    GoTo changeFound
  End If
  
  FaxNum = ReplaceString(fptxtCntctFax.Text, "(", "")
  FaxNum = ReplaceString(FaxNum, ")", "")
  FaxNum = ReplaceString(FaxNum, "-", "")
  If QPTrim$(FaxNum) <> QPTrim$(W2RARec.CntFAX) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fptxtCntctFax.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmbPrefMeth.Text, 1, 1) <> QPTrim$(W2RARec.CntMethod) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbPrefMeth.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmbPrepCode.Text, 1, 1) <> QPTrim$(W2RARec.PrepCode) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 0
    fpcmbPrepCode.SetFocus
    GoTo changeFound
  End If
  
NoRAFile:
  If Exist("PRDATA\W2ESUBRE.DAT") Then
    OpenW2ESubRE REHandle
    Get REHandle, 1, W2RERec
    Close REHandle
  Else
    GoTo NoFile
  End If
  
  If QPTrim$(fptxtTaxYear.Text) <> QPTrim$(W2RERec.TaxYear) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtTaxYear.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmbAgentCode.Text, 1, 1) <> QPTrim$(W2RERec.AgentCode) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fpcmbAgentCode.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtAgent4EIN.Text) <> QPTrim$(W2RERec.EINAgent) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtAgent4EIN.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmprAgtEIN.Text) <> QPTrim$(W2RERec.EmprAgntEIN) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmprAgtEIN.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtOtherEIN.Text) <> QPTrim$(W2RERec.OthEIN) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtOtherEIN.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmbTermBusInd.Text, 1, 1) <> QPTrim$(W2RERec.TermBusInd) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fpcmbTermBusInd.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmployerName.Text) <> QPTrim$(W2RERec.EmprName) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmployerName.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmprAdd1.Text) <> QPTrim$(W2RERec.EmprDelAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmprAdd1.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmprAdd2.Text) <> QPTrim$(W2RERec.EmprLocAddr) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmprAdd2.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmprCity.Text) <> QPTrim$(W2RERec.EmprCity) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmprCity.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fpcmbEmprState.Text) <> QPTrim$(W2RERec.EmprState) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fpcmbEmprState.SetFocus
    GoTo changeFound
  End If
  
  If QPTrim$(fptxtEmprZip.Text) <> QPTrim$(W2RERec.EmprZip) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fptxtEmprZip.SetFocus
    GoTo changeFound
  End If
  
  If Mid(fpcmb3rdSckPay.Text, 1, 1) <> QPTrim$(W2RERec.ThrdSckPay) Then
    changeFlag = True
    vaTabPro1.ActiveTab = 1
    fpcmb3rdSckPay.SetFocus
    GoTo changeFound
  End If
  
changeFound:
  If changeFlag = True Then
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      frmW2Processing.Show
      DoEvents
      Unload frmW2EmpInfo
    Case Else:
       'Do nothing because we don't know about any options except
       'save, review or abandon...used as a placeholder for adding
       'other options at a later date
    End Select
  Else
    changeFlag = False
    frmW2Processing.Show
    DoEvents
    Unload frmW2EmpInfo
  End If
NoFile:
  frmW2Processing.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim W2RARec As W2ElectronicSubRA
  Dim RAHandle As Integer
  Dim W2RERec As W2ElectronicSubRE
  Dim REHandle As Integer
  Dim FaxNum$, PhoneNum$
  Dim W2RWRec As W2ElectronicSubRW
  Dim W2SetUpRec As W2SetUpType
  Dim W2Handle As Integer
  Dim W2RSRec As W2ElectronicSubRS
  Dim W2RSLen As Integer
  Dim TaxYear$
  Dim tStr$
  
  If Len(QPTrim$(fptxtPinNum.Text)) <> 8 Then 'added 8/25/04 because accuwage refused any number
    frmW2MessageWOpts.Label1.Caption = "According to IRS regulations the Personal Identification Number (PIN) must be 8 alphanumeric characters. The number entered will cause an error if Accuwage tries to process it. Press F10 to continue anyway or press ESC to stop this process now."
    frmW2MessageWOpts.Label1.Top = 700
    frmW2MessageWOpts.Show vbModal
    If frmW2MessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmW2MessageWOpts
      vaTabPro1.SetFocus
      vaTabPro1.ActiveTab = 0
      fptxtPinNum.SetFocus
      Close
      Exit Sub
    Else
      Unload frmW2MessageWOpts
    End If
  End If
      
  If Exist(W2SetupFile) Then 'added 8/25/04
    OpenW2SetUp W2Handle
    Get W2Handle, 1, W2SetUpRec
    Close W2Handle
    TaxYear = CStr(W2SetUpRec.ExtrYear)
  Else
    TaxYear = "0000"
  End If
  
  If QPTrim$(fptxtTaxYear.Text) <> TaxYear Then 'added 8/25/04
    frmW2MessageWOpts.Label1.Caption = "You are attempting to process tax data for " + fptxtTaxYear.Text + " but the currentt year for which data has been extracted is " + TaxYear + ". Press F10 to continue anyway or press ESC to stop processing now."
    frmW2MessageWOpts.Label1.Top = 700
    frmW2MessageWOpts.Show vbModal
    If frmW2MessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmW2MessageWOpts
      vaTabPro1.SetFocus
      vaTabPro1.ActiveTab = 1
      fptxtTaxYear.SetFocus
      Close
      Exit Sub
    Else
      Unload frmW2MessageWOpts
    End If
  End If
  
  If QPTrim$(fptxtEmpEIN.Text) = "" Then
    MsgBox "Please enter a Submitter's Employer Identification Number (EIN)."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtEmpEIN.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtPinNum.Text) = "" Then
    MsgBox "Please enter a Personal Identification Number (PIN)."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtPinNum.SetFocus
    Exit Sub
  End If
  
  If Mid(fpcmbResub.Text, 1, 1) = "" Then
    MsgBox "Please enter a Resub ID."
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbResub.SetFocus
    Exit Sub
  End If
  
  If Mid(fpcmbResub.Text, 1, 1) = "1" Then
    If fptxtResubWFID.Text = "" Then
      MsgBox "Please enter a Resub WFID number."
      Close
      vaTabPro1.ActiveTab = 0
      fptxtResubWFID.SetFocus
      Exit Sub
    ElseIf Len(fptxtResubWFID.Text) <> 6 Then
      MsgBox "The Resub WFID number must be 6 digits long."
      Close
      vaTabPro1.ActiveTab = 0
      fptxtResubWFID.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fptxtTownName.Text) = "" Then
    MsgBox "Please enter a Town Name."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtTownName.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdd1.Text) = "" And QPTrim$(fptxtAdd2.Text) = "" Then
    MsgBox "Please enter either the town's Delivery Address or Location Address."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    MsgBox "Please enter a City name."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbState.Text) = "" Then
    MsgBox "Please enter the town's State."
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtZip.Text) = "" Then
    MsgBox "Please enter the town's Zip Code."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If Len(fptxtZip.Text) <> 5 Then
    MsgBox "Please enter a 5 digit number for the town's Zip Code."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtZipX.Text) <> "" Then
    If Len(fptxtZipX.Text) <> 4 Then
      MsgBox "Please enter a 4 digit number for the town's Zip Code extension. Otherwise leave it blank."
      Close
      vaTabPro1.ActiveTab = 0
      fptxtZipX.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fptxtSubmitterName.Text) = "" Then
    MsgBox "Please enter a Submitter Name."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtSubmitterName.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtSubAdd1.Text) = "" And QPTrim$(fptxtSubAdd2.Text) = "" Then
    MsgBox "Please enter either the submitter's Delivery Address or Location Address."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtSubAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtSubCity.Text) = "" Then
    MsgBox "Please enter the submitter's City."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtSubCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbSubState.Text) = "" Then
    MsgBox "Please enter the submitter's State."
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbSubState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtSubZip.Text) = "" Then
    MsgBox "Please enter the submitter's Zip Code."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtSubZip.SetFocus
    Exit Sub
  End If
  
  If Len(fptxtSubZip.Text) <> 5 Then
    MsgBox "Please enter a 5 digit number for the submitter's Zip Code. Otherwise leave it blank."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtSubZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtSubZipX.Text) <> "" Then
    If Len(fptxtSubZipX.Text) <> 4 Then
      MsgBox "Please enter a 4 digit number for the submitter's Zip Code extension."
      Close
      vaTabPro1.ActiveTab = 0
      fptxtSubZip.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fptxtCntctName.Text) = "" Then
    MsgBox "Please enter a Contact Name "
    Close
    vaTabPro1.ActiveTab = 0
    fptxtCntctName.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtPhone.Text) = "(" Then
    MsgBox "Please enter a contact's Phone Number."
    Close
    vaTabPro1.ActiveTab = 0
    fptxtPhone.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbPrefMeth.Text) = "" Then
    MsgBox "Please enter the Preferred Method of Problem Notification Code."
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbPrefMeth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbPrepCode.Text) = "" Then
    MsgBox "Please enter a Preparer Code selection."
    Close
    vaTabPro1.ActiveTab = 0
    fpcmbPrepCode.SetFocus
    Exit Sub
  End If
  
  If Mid(fpcmbPrefMeth.Text, 1, 1) = "1" Then
    If QPTrim$(fptxtCntctEMail.Text) = "" Then
      MsgBox "Since the preferred method of contact is through email then please enter the Contact's Email Address. "
      Close
      vaTabPro1.ActiveTab = 0
      fptxtCntctEMail.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fptxtTaxYear.Text) = "" Then
    MsgBox "Please enter the Tax Year."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtTaxYear.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbTermBusInd.Text) = "" Then
    MsgBox "A terminated business code is required."
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbTermBusInd.SetFocus
    Exit Sub
  End If
  
  If Mid(fpcmbAgentCode.Text, 1, 1) = "1" Then
    If QPTrim$(fptxtAgent4EIN.Text) = "" Then
      MsgBox "Since 1 is selected as the Agent Indicator Code please enter an employer's EIN number for the agent."
      Close
      vaTabPro1.ActiveTab = 1
      fptxtAgent4EIN.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fptxtEmprAgtEIN.Text) = "" Then
    MsgBox "Please enter the EIN entered on the form 941 submitted to the IRS."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtEmprAgtEIN.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtEmployerName.Text) = "" Then
    MsgBox "Please enter an Employer Name."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtEmployerName.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtEmprAdd1.Text) = "" And QPTrim$(fptxtEmprAdd2.Text) = "" Then
    MsgBox "Please enter either a Delivery Address or a Location Address for the Employer."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtEmprAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtEmprCity.Text) = "" Then
    MsgBox "Please enter the Employer's City."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtEmprCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbEmprState.Text) = "" Then
    MsgBox "Please enter the Employer's State."
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbEmprState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtEmprZip.Text) = "" Then
    MsgBox "Please enter the Employer's Zip Code."
    Close
    vaTabPro1.ActiveTab = 1
    fptxtEmprZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmb3rdSckPay.Text) = "" Then
    MsgBox "Please enter the Third Party Sick Pay Indicator."
    Close
    vaTabPro1.ActiveTab = 1
    fpcmb3rdSckPay.SetFocus
    Exit Sub
  End If
  
  tStr$ = Mid$(fpcmbEmpKind.Text, 1, 1)
  If Len(QPTrim$(tStr$)) = 0 Then
    MsgBox "Please enter the Kind of Employer Indicator."
    Close
    vaTabPro1.ActiveTab = 1
    fpcmbEmpKind.SetFocus
    Exit Sub
   End If
    
  OpenW2ESubRA RAHandle
  
  W2RARec.EINNum = QPTrim$(fptxtEmpEIN.Text)
  W2RARec.PersIDNum = QPTrim$(fptxtPinNum.Text)
  W2RARec.ResubID = Mid(fpcmbResub.Text, 1, 1)
  W2RARec.ReSubWFID = QPTrim$(fptxtResubWFID.Text)
  W2RARec.SftwrCode = "99"
  W2RARec.CmpnyName = QPTrim$(fptxtTownName.Text)
  W2RARec.DelAddr = QPTrim$(fptxtAdd1.Text)
  W2RARec.LocAddr = QPTrim$(fptxtAdd2.Text)
  W2RARec.City = QPTrim$(fptxtCity.Text)
  W2RARec.State = QPTrim$(fpcmbState.Text)
  W2RARec.Zip = QPTrim$(fptxtZip.Text)
  W2RARec.ZipExt = QPTrim$(fptxtZipX.Text)
  W2RARec.SubmttrName = QPTrim$(fptxtSubmitterName.Text)
  W2RARec.SubDelAddr = QPTrim$(fptxtSubAdd1.Text)
  W2RARec.SubLocAddr = QPTrim$(fptxtSubAdd2.Text)
  W2RARec.SubCity = QPTrim$(fptxtSubCity.Text)
  W2RARec.SubState = QPTrim$(fpcmbSubState.Text)
  W2RARec.SubZip = QPTrim$(fptxtSubZip.Text)
  W2RARec.SubZipExt = QPTrim$(fptxtSubZipX.Text)
  W2RARec.ContactName = QPTrim$(fptxtCntctName.Text)
  PhoneNum = ReplaceString(fptxtPhone.Text, "(", "")
  PhoneNum = ReplaceString(PhoneNum, ")", "")
  PhoneNum = ReplaceString(PhoneNum, "-", "")
  W2RARec.CntctPhone = QPTrim(PhoneNum)
  W2RARec.CntPhnExt = QPTrim$(fptxtPhoneX.Text)
  W2RARec.CntEMail = QPTrim$(fptxtCntctEMail.Text)
  FaxNum = ReplaceString(fptxtCntctFax.Text, "(", "")
  FaxNum = ReplaceString(FaxNum, ")", "")
  FaxNum = ReplaceString(FaxNum, "-", "")
  W2RARec.CntFAX = QPTrim$(FaxNum)
  W2RARec.CntMethod = QPTrim$(fpcmbPrefMeth.Text)
  W2RARec.PrepCode = QPTrim$(fpcmbPrepCode.Text)
  Put RAHandle, 1, W2RARec
  Close RAHandle

  OpenW2ESubRE REHandle
  
  W2RERec.TaxYear = fptxtTaxYear
  W2RERec.AgentCode = Mid(fpcmbAgentCode.Text, 1, 1)
  W2RERec.EINAgent = QPTrim$(fptxtAgent4EIN.Text)
  W2RERec.EmprAgntEIN = QPTrim$(fptxtEmprAgtEIN.Text)
  W2RERec.OthEIN = QPTrim$(fptxtOtherEIN.Text)
  W2RERec.TermBusInd = Mid(fpcmbTermBusInd.Text, 1, 1)
  W2RERec.EmprName = QPTrim$(fptxtEmployerName.Text)
  W2RERec.EmprDelAddr = QPTrim$(fptxtEmprAdd1.Text)
  W2RERec.EmprLocAddr = QPTrim$(fptxtEmprAdd2.Text)
  W2RERec.EmprCity = QPTrim$(fptxtEmprCity.Text)
  W2RERec.EmprState = QPTrim$(fpcmbEmprState.Text)
  W2RERec.EmprZip = QPTrim$(fptxtEmprZip.Text)
  W2RERec.EmprZipX = QPTrim$(fptxtEmprZipX.Text)
  W2RERec.ThrdSckPay = QPTrim$(fpcmb3rdSckPay.Text)
  W2RERec.EmprType = tStr$
  Put REHandle, 1, W2RERec
  Close REHandle
  
  If MsgBox("The RA and RE records required for W2 Electronic Submission File have been saved successfully. Do you wish to build the W2 Electronic Submission File now?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  Call BuildEFile
  If Exist("PRDATA\W2ESUBRW.DAT") Then
    MsgBox "Your W2 Electronic Submission File has been successfully created inside the 'W2REPORT' directory located in your Citipak Directory."
  End If
  
  frmW2Processing.Show
  DoEvents
  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
    If vaTabPro1.ActiveTab = 0 Then
      vaTabPro1.ActiveTab = 1
      fptxtTaxYear.SetFocus
    ElseIf vaTabPro1.ActiveTab = 1 Then
      vaTabPro1.ActiveTab = 0
      fptxtEmpEIN.SetFocus
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadThisForm
  Me.HelpContextID = hlpW2Electronic
  vaTabPro1.ActiveTab = 1
  DoEvents
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
''    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmMedW2Setup.")
      End
    End If
  End If
End Sub

Public Sub LoadThisForm()
  
  Dim UHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim x As Integer
  Dim W2RARec As W2ElectronicSubRA
  Dim RAHandle As Integer
  Dim W2RERec As W2ElectronicSubRE
  Dim REHandle As Integer
  Dim tStr As String
'
  OpenUnitFile UHandle
  If LOF(UHandle) / Len(UnitRec) = 0 Then
    MsgBox "Please save employer information on the Employer File screen before continuing."
    Close
    Exit Sub
  End If
  Get UHandle, 1, UnitRec
  Close UHandle
  
  If Exist("PRDATA\W2ESUBRA.DAT") Then
    OpenW2ESubRA RAHandle
    Get RAHandle, 1, W2RARec
    Close RAHandle
    
    fptxtEmpEIN.Text = QPTrim$(W2RARec.EINNum)
    fptxtPinNum.Text = QPTrim$(W2RARec.PersIDNum)
    If W2RARec.ResubID = "1" Then
      fpcmbResub.Text = "1 Resubmission"
    Else
      fpcmbResub.Text = "0 Not A Resub"
    End If
      
    If QPTrim$(W2RARec.CmpnyName) = "" And QPTrim$(UnitRec.UFEMPR) = "" Then
      fptxtTownName.Text = ""
    ElseIf QPTrim$(W2RARec.CmpnyName) <> "" Then
      fptxtTownName.Text = QPTrim$(W2RARec.CmpnyName)
    Else
      fptxtTownName.Text = QPTrim$(UnitRec.UFEMPR)
    End If
    If QPTrim$(W2RARec.DelAddr) = "" And QPTrim$(UnitRec.UFADDR1) = "" Then
      fptxtAdd1.Text = ""
    ElseIf QPTrim$(W2RARec.DelAddr) <> "" Then
      fptxtAdd1.Text = QPTrim$(W2RARec.DelAddr)
    Else
      fptxtAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    End If
    
    If QPTrim$(W2RARec.LocAddr) = "" And QPTrim$(UnitRec.UFADDR2) = "" Then
      fptxtAdd2.Text = ""
    ElseIf QPTrim$(W2RARec.LocAddr) <> "" Then
      fptxtAdd2.Text = QPTrim$(W2RARec.LocAddr)
    Else
      fptxtAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    End If
    If QPTrim$(W2RARec.City) = "" And QPTrim$(UnitRec.UFCITY) = "" Then
      fptxtCity.Text = ""
    ElseIf QPTrim$(W2RARec.City) <> "" Then
      fptxtCity.Text = QPTrim$(W2RARec.City)
    Else
      fptxtCity.Text = QPTrim$(UnitRec.UFCITY)
    End If
    
    If QPTrim$(W2RARec.State) = "" And QPTrim$(UnitRec.UFSTATE) = "" Then
      fpcmbState.Text = ""
    ElseIf QPTrim$(W2RARec.State) <> "" Then
      fpcmbState.Text = QPTrim$(W2RARec.State)
    Else
      fpcmbState.Text = QPTrim$(UnitRec.UFSTATE)
    End If
    If QPTrim$(W2RARec.Zip) = "" And Mid(UnitRec.UFZIP, 1, 5) = "" Then
      fptxtZip.Text = ""
    ElseIf QPTrim$(W2RARec.Zip) <> "" Then
      fptxtZip.Text = QPTrim$(W2RARec.Zip)
    Else
      fptxtZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    End If
    
    If QPTrim$(W2RARec.ZipExt) = "" And QPTrim$(Mid(UnitRec.UFZIP, 7, 4)) = "" Then
      fptxtZipX.Text = ""
    ElseIf QPTrim$(W2RARec.ZipExt) <> "" Then
      fptxtZipX.Text = QPTrim$(W2RARec.ZipExt)
    Else
      fptxtZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    End If
    If QPTrim$(W2RARec.SubmttrName) = "" And QPTrim$(UnitRec.UFATTN) = "" Then
      fptxtSubmitterName.Text = ""
    ElseIf QPTrim$(W2RARec.SubmttrName) <> "" Then
      fptxtSubmitterName.Text = QPTrim$(W2RARec.SubmttrName)
    Else
      fptxtSubmitterName.Text = QPTrim$(UnitRec.UFATTN)
    End If
    
    If QPTrim$(W2RARec.SubDelAddr) = "" And QPTrim$(UnitRec.UFADDR1) = "" Then
      fptxtSubAdd1.Text = ""
    ElseIf QPTrim$(W2RARec.SubDelAddr) <> "" Then
      fptxtSubAdd1.Text = QPTrim$(W2RARec.SubDelAddr)
    Else
      fptxtSubAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    End If
    If QPTrim$(W2RARec.SubLocAddr) = "" And QPTrim$(UnitRec.UFADDR2) = "" Then
      fptxtSubAdd2.Text = ""
    ElseIf QPTrim$(W2RARec.SubLocAddr) <> "" Then
      fptxtSubAdd2.Text = QPTrim$(W2RARec.SubLocAddr)
    Else
      fptxtSubAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    End If
    
    If QPTrim$(W2RARec.SubCity) = "" And QPTrim$(UnitRec.UFEMPR) = "" Then
      fptxtSubCity.Text = ""
    ElseIf QPTrim$(W2RARec.SubCity) <> "" Then
      fptxtSubCity.Text = QPTrim$(W2RARec.SubCity)
    Else
      fptxtSubCity.Text = QPTrim$(UnitRec.UFEMPR)
    End If
    If QPTrim$(W2RARec.SubState) = "" And QPTrim$(UnitRec.UFSTATE) = "" Then
      fpcmbSubState.Text = ""
    ElseIf QPTrim$(W2RARec.SubState) <> "" Then
      fpcmbSubState.Text = QPTrim$(W2RARec.SubState)
    Else
      fpcmbSubState.Text = QPTrim$(UnitRec.UFSTATE)
    End If
    
    If QPTrim$(W2RARec.SubZip) = "" And Mid(UnitRec.UFZIP, 1, 5) = "" Then
      fptxtSubZip.Text = ""
    ElseIf QPTrim$(W2RARec.SubZip) <> "" Then
      fptxtSubZip.Text = QPTrim$(W2RARec.SubZip)
    Else
      fptxtSubZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    End If
    If QPTrim$(W2RARec.SubZipExt) = "" And QPTrim$(Mid(UnitRec.UFZIP, 7, 4)) = "" Then
      fptxtSubZipX.Text = ""
    ElseIf QPTrim$(W2RARec.SubZipExt) <> "" Then
      fptxtSubZipX.Text = QPTrim$(W2RARec.SubZipExt)
    Else
      fptxtSubZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    End If
    
    If QPTrim$(W2RARec.ContactName) = "" And QPTrim$(UnitRec.UFATTN) = "" Then
      fptxtCntctName.Text = ""
    ElseIf QPTrim$(W2RARec.ContactName) <> "" Then
      fptxtCntctName.Text = QPTrim$(W2RARec.ContactName)
    Else
      fptxtCntctName.Text = QPTrim$(UnitRec.UFATTN)
    End If
    
    fptxtPhone.Text = QPTrim$(W2RARec.CntctPhone)
    fptxtPhoneX.Text = QPTrim$(W2RARec.CntPhnExt)
    fptxtCntctEMail.Text = QPTrim$(W2RARec.CntEMail)
    fptxtCntctFax.Text = QPTrim$(W2RARec.CntFAX)
    If QPTrim$(W2RARec.CntMethod) = "1" Then
      fpcmbPrefMeth.Text = "1 EMail/Internet"
    ElseIf QPTrim$(W2RARec.CntMethod) = "2" Then
      fpcmbPrefMeth.Text = "2 Postal Service"
    End If
    If QPTrim$(W2RARec.PrepCode) = "A" Then
      fpcmbPrepCode.Text = "A Accounting Firm"
    ElseIf QPTrim$(W2RARec.PrepCode) = "L" Then
      fpcmbPrepCode.Text = "L Self-Prepared"
    ElseIf QPTrim$(W2RARec.PrepCode) = "S" Then
      fpcmbPrepCode.Text = "S Service Bureau"
    ElseIf QPTrim$(W2RARec.PrepCode) = "P" Then
      fpcmbPrepCode.Text = "P Parent Company"
    ElseIf QPTrim$(W2RARec.PrepCode) = "O" Then
      fpcmbPrepCode.Text = "O Other"
    End If
  Else '---------------------------------------------
    fptxtEmpEIN.Text = ""
    fptxtPinNum.Text = ""
    fpcmbResub.Text = "0 Not A Resub"
    fptxtResubWFID = ""
    If QPTrim$(UnitRec.UFEMPR) = "" Then
      fptxtTownName.Text = ""
    Else
      fptxtTownName.Text = QPTrim$(UnitRec.UFEMPR)
    End If
    If QPTrim$(UnitRec.UFADDR1) = "" Then
      fptxtAdd1.Text = ""
    Else
      fptxtAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    End If
    
    If QPTrim$(UnitRec.UFADDR2) = "" Then
      fptxtAdd2.Text = ""
    Else
      fptxtAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    End If
    If QPTrim$(UnitRec.UFCITY) = "" Then
      fptxtCity.Text = ""
    Else
      fptxtCity.Text = QPTrim$(UnitRec.UFCITY)
    End If
    
    If QPTrim$(UnitRec.UFSTATE) = "" Then
      fpcmbState.Text = ""
    Else
      fpcmbState.Text = QPTrim$(UnitRec.UFSTATE)
    End If
    If Mid(UnitRec.UFZIP, 1, 5) = "" Then
      fptxtZip.Text = ""
    Else
      fptxtZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    End If
    
    If QPTrim$(Mid(UnitRec.UFZIP, 6, 4)) = "" Then
      fptxtZipX.Text = ""
    Else
      fptxtZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    End If
    If QPTrim$(UnitRec.UFATTN) = "" Then
      fptxtSubmitterName.Text = ""
    Else
      fptxtSubmitterName.Text = QPTrim$(UnitRec.UFATTN)
    End If
    
    If QPTrim$(UnitRec.UFADDR1) = "" Then
      fptxtSubAdd1.Text = ""
    Else
      fptxtSubAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    End If
    If QPTrim$(UnitRec.UFADDR2) = "" Then
      fptxtSubAdd2.Text = ""
    Else
      fptxtSubAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    End If
    
    If QPTrim$(UnitRec.UFCITY) = "" Then
      fptxtSubCity.Text = ""
    Else
      fptxtSubCity.Text = QPTrim$(UnitRec.UFCITY)
    End If
    If QPTrim$(UnitRec.UFSTATE) = "" Then
      fpcmbSubState.Text = ""
    Else
      fpcmbSubState.Text = QPTrim$(UnitRec.UFSTATE)
    End If
    
    If Mid(UnitRec.UFZIP, 1, 5) = "" Then
      fptxtSubZip.Text = ""
    Else
      fptxtSubZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    End If
    If QPTrim$(Mid(UnitRec.UFZIP, 6, 4)) = "" Then
      fptxtSubZipX.Text = ""
    Else
      fptxtSubZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    End If
    
    If QPTrim$(UnitRec.UFATTN) = "" Then
      fptxtCntctName.Text = ""
    Else
      fptxtCntctName.Text = QPTrim$(UnitRec.UFATTN)
    End If
    
    fptxtPhone.Text = ""
    fptxtPhoneX.Text = ""
    fptxtCntctEMail.Text = ""
    fptxtCntctFax.Text = ""
    fpcmbPrefMeth.Text = ""
    fpcmbPrepCode.Text = ""
  End If
  
  If Exist("PRDATA\W2ESUBRE.DAT") Then
    OpenW2ESubRE REHandle
    Get REHandle, 1, W2RERec
    Close REHandle
    
    fptxtTaxYear = W2RERec.TaxYear
    If W2RERec.AgentCode = "1" Then
      fpcmbAgentCode.Text = "1 2678 Agent(Approved by IRS)"
    ElseIf W2RERec.AgentCode = "2" Then
      fpcmbAgentCode.Text = "2 Common Pay Master"
    End If
    fptxtAgent4EIN.Text = QPTrim$(W2RERec.EINAgent)
    fptxtEmprAgtEIN.Text = QPTrim$(W2RERec.EmprAgntEIN)
    fptxtOtherEIN.Text = QPTrim$(W2RERec.OthEIN)
    If W2RERec.TermBusInd = "0" Then
      fpcmbTermBusInd.Text = "0 Not Terminated"
    ElseIf W2RERec.TermBusInd = "1" Then
      fpcmbTermBusInd.Text = "1 Terminated This Year"
    End If
    
    If QPTrim$(W2RERec.EmprName) <> "" Then
      fptxtEmployerName.Text = QPTrim$(W2RERec.EmprName)
    ElseIf QPTrim$(UnitRec.UFEMPR) <> "" Then
      fptxtEmployerName.Text = QPTrim$(UnitRec.UFEMPR)
    Else
      fptxtEmployerName.Text = ""
    End If
    If QPTrim$(W2RERec.EmprDelAddr) <> "" Then
      fptxtEmprAdd1.Text = QPTrim$(W2RERec.EmprDelAddr)
    ElseIf QPTrim$(UnitRec.UFADDR1) <> "" Then
      fptxtEmprAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    Else
      fptxtEmprAdd1.Text = ""
    End If
    
    If QPTrim$(W2RERec.EmprLocAddr) <> "" Then
      fptxtEmprAdd2.Text = QPTrim$(W2RERec.EmprLocAddr)
    ElseIf QPTrim$(UnitRec.UFADDR2) <> "" Then
      fptxtEmprAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    Else
      fptxtEmprAdd2.Text = ""
    End If
    If QPTrim$(W2RERec.EmprCity) <> "" Then
      fptxtEmprCity.Text = QPTrim$(W2RERec.EmprCity)
    ElseIf QPTrim$(UnitRec.UFCITY) <> "" Then
      fptxtEmprCity.Text = QPTrim$(UnitRec.UFCITY)
    Else
      fptxtEmprCity.Text = ""
    End If
    
    If QPTrim$(W2RERec.EmprState) <> "" Then
      fpcmbEmprState.Text = QPTrim$(W2RERec.EmprState)
    ElseIf QPTrim$(UnitRec.UFSTATE) <> "" Then
      fpcmbEmprState.Text = QPTrim$(UnitRec.UFSTATE)
    Else
      fpcmbEmprState.Text = ""
    End If
    If QPTrim$(W2RERec.EmprZip) <> "" Then
      fptxtEmprZip.Text = QPTrim$(W2RERec.EmprZip)
    ElseIf Mid(UnitRec.UFZIP, 1, 5) <> "" Then
      fptxtEmprZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    Else
      fptxtEmprZip.Text = ""
    End If
    
    If QPTrim$(W2RERec.EmprZipX) <> "" Then
      fptxtEmprZipX.Text = QPTrim$(W2RERec.EmprZipX)
    ElseIf QPTrim$(Mid(UnitRec.UFZIP, 6, 4)) <> "" Then
      fptxtEmprZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    Else
      fptxtEmprZipX.Text = ""
    End If
    
    If W2RERec.ThrdSckPay = "0" Then
      fpcmb3rdSckPay.Text = "0 No Sick Pay"
    ElseIf W2RERec.ThrdSckPay = "1" Then
      fpcmb3rdSckPay.Text = "1 Sick Pay"
    End If
    
    Select Case W2RERec.EmprType
    Case "F"
      tStr$ = "F  Federal Government"
    Case "S"
      tStr$ = "S  State/Local Governmental Employer"
    Case "T"
      tStr$ = "T  Tax Exempt Employer"
    Case "Y"
      tStr$ = "Y  State/Local Tax Exempt Employer"
    Case "N"
      tStr$ = "N  None Apply"
    End Select
    fpcmbEmpKind.Text = tStr$
    
'    Dale this
    
  Else '----------------------------------------------
    fptxtTaxYear = Mid(Date, 7, 4)
    fpcmbAgentCode.Text = ""
    fptxtAgent4EIN.Text = ""
    fptxtEmprAgtEIN.Text = ""
    fptxtOtherEIN.Text = ""
    fpcmbTermBusInd.Text = ""
    
    If QPTrim$(UnitRec.UFEMPR) <> "" Then
      fptxtEmployerName.Text = QPTrim$(UnitRec.UFEMPR)
    Else
      fptxtEmployerName.Text = ""
    End If
    If QPTrim$(UnitRec.UFADDR1) <> "" Then
      fptxtEmprAdd1.Text = QPTrim$(UnitRec.UFADDR1)
    Else
      fptxtEmprAdd1.Text = ""
    End If
    
    If QPTrim$(UnitRec.UFADDR2) <> "" Then
      fptxtEmprAdd2.Text = QPTrim$(UnitRec.UFADDR2)
    Else
      fptxtEmprAdd2.Text = ""
    End If
    If QPTrim$(UnitRec.UFCITY) <> "" Then
      fptxtEmprCity.Text = QPTrim$(UnitRec.UFCITY)
    Else
      fptxtEmprCity.Text = ""
    End If
    
    If QPTrim$(UnitRec.UFSTATE) <> "" Then
      fpcmbEmprState.Text = QPTrim$(UnitRec.UFSTATE)
    Else
      fpcmbEmprState.Text = ""
    End If
    If Mid(UnitRec.UFZIP, 1, 5) <> "" Then
      fptxtEmprZip.Text = Mid(UnitRec.UFZIP, 1, 5)
    Else
      fptxtEmprZip.Text = ""
    End If
    
    If QPTrim$(Mid(UnitRec.UFZIP, 6, 4)) <> "" Then
      fptxtEmprZipX.Text = Mid(UnitRec.UFZIP, 6, 4)
    Else
      fptxtEmprZipX.Text = ""
    End If
    
    fpcmb3rdSckPay.Text = ""
  End If
    
  fpcmbResub.AddItem "1 Resubmission"
  fpcmbResub.AddItem "0 Not A Resub"
    
  fpcmbPrefMeth.AddItem "1 EMail/Internet"
  fpcmbPrefMeth.AddItem "2 Postal Service"
  fpcmbPrepCode.AddItem "A Accounting Firm"
  fpcmbPrepCode.AddItem "L Self-Prepared"
  fpcmbPrepCode.AddItem "S Service Bureau"
  fpcmbPrepCode.AddItem "P Parent Company"
  fpcmbPrepCode.AddItem "O Other"
  fpcmbState.AddItem "AL"
  fpcmbState.AddItem "AR"
  fpcmbState.AddItem "GA"
  fpcmbState.AddItem "NC"
  fpcmbState.AddItem "SC"
  fpcmbState.AddItem "TN"
  fpcmbState.AddItem "VA"
  fpcmbSubState.AddItem "AL"
  fpcmbSubState.AddItem "AR"
  fpcmbSubState.AddItem "GA"
  fpcmbSubState.AddItem "NC"
  fpcmbSubState.AddItem "SC"
  fpcmbSubState.AddItem "TN"
  fpcmbSubState.AddItem "VA"
  
  fpcmbEmpKind.AddItem "F  Federal Government"
  fpcmbEmpKind.AddItem "S  State/Local Governmental Employer"
  fpcmbEmpKind.AddItem "T  Tax Exempt Employer"
  fpcmbEmpKind.AddItem "Y  State/Local Tax Exempt Employer"
  fpcmbEmpKind.AddItem "N  None Apply"

  fpcmbAgentCode.AddItem "1 2678 Agent (Approved by IRS)"
  fpcmbAgentCode.AddItem "2 Common Pay Master"
  fpcmbTermBusInd.AddItem "0 Not Terminated"
  fpcmbTermBusInd.AddItem "1 Terminated This Year"
  fpcmbEmprState.AddItem "AL"
  fpcmbEmprState.AddItem "AR"
  fpcmbEmprState.AddItem "GA"
  fpcmbEmprState.AddItem "NC"
  fpcmbEmprState.AddItem "SC"
  fpcmbEmprState.AddItem "TN"
  fpcmbEmprState.AddItem "VA"
  fpcmb3rdSckPay.AddItem "0 No Sick Pay"
  fpcmb3rdSckPay.AddItem "1 Sick Pay"
  
End Sub
Private Sub fpcmb3rdSckPay_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmb3rdSckPay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmb3rdSckPay.ListIndex = -1
  End If
  If fpcmb3rdSckPay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtTaxYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmb3rdSckPay_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub fpcmbAgentCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbAgentCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAgentCode.ListIndex = -1
  End If
  If fpcmbAgentCode.ListDown <> True Then
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

Private Sub fpcmbAgentCode_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub


Private Sub fpcmbEmpKind_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
  Case vbKeyPageDown, vbKeyPageUp
    KeyCode = 0
  Case vbKeySpace
    fpcmbEmpKind.ListDown = True
 
  End Select
  
End Sub

Private Sub fpcmbEmprState_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEmprState.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEmprState.ListIndex = -1
  End If
  If fpcmbEmprState.ListDown <> True Then
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

Private Sub fpcmbEmprState_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcmbPrefMeth_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrefMeth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrefMeth.ListIndex = -1
  End If
  If fpcmbPrefMeth.ListDown <> True Then
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

Private Sub fpcmbPrefMeth_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcmbPrepCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrepCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrepCode.ListIndex = -1
  End If
  If fpcmbPrepCode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtEmpEIN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbResub_Change()
  If fpcmbResub.Text = "1" Then
    fptxtResubWFID.Enabled = True
  Else
    fptxtResubWFID.Enabled = False
  End If
End Sub

Private Sub fpcmbResub_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbResub.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbResub.ListIndex = -1
  End If
  If fpcmbResub.ListDown <> True Then
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

Private Sub fpcmbState_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbState.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbState.ListIndex = -1
  End If
  If fpcmbState.ListDown <> True Then
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

Private Sub fpcmbState_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcmbSubState_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSubState.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSubState.ListIndex = -1
  End If
  If fpcmbSubState.ListDown <> True Then
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

Private Sub fpcmbSubState_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcmbTermBusInd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTermBusInd.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTermBusInd.ListIndex = -1
  End If
  If fpcmbTermBusInd.ListDown <> True Then
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

Private Sub fpcmbTermBusInd_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtAdd1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtAdd2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtAgent4EIN_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtCity_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtCntctFax_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtCntctName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmpEIN_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmployerName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprAdd1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprAdd2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprAgtEIN_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprCity_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprZip_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtEmprZipX_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtOtherEIN_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtPhone_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtPhoneX_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtPinNum_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubAdd1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubAdd2_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubCity_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubmitterName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubZip_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtSubZipX_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtTaxYear_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtTownName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtZip_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtZipX_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub vaTabPro1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
    If vaTabPro1.ActiveTab = 0 Then
      vaTabPro1.ActiveTab = 1
      fpcmbEmpKind.SetFocus
      'fptxtTaxYear.SetFocus
    ElseIf vaTabPro1.ActiveTab = 1 Then
      vaTabPro1.ActiveTab = 0
      fptxtEmpEIN.SetFocus
    End If
  End If
End Sub

Private Sub vaTabPro1_TabShown(ActiveTab As Integer)
'  If ActiveTab = 1 Then
'    fptxtTaxYear.SetFocus
'  ElseIf ActiveTab = 0 Then
'    fptxtEmpEIN.SetFocus
'  End If
End Sub

Private Sub BuildEFile()
  Dim W2RARec As W2ElectronicSubRA
  Dim RAHandle As Integer
  Dim W2RERec As W2ElectronicSubRE
  Dim REHandle As Integer
  Dim W2RWRec As W2ElectronicSubRW
  Dim RWHandle As Integer
  Dim NumOfRWRecs As Integer
  Dim W2RTRec As W2ElectronicSubRT
  Dim RTHandle As Integer
  Dim W2RFRec As W2ElectronicSubRF
  Dim RFHandle As Integer
  Dim W2RORec As W2ElectronicSubRO
  Dim ROHandle As Integer
  Dim W2RURec As W2ElectronicSubRU
  Dim RUHandle As Integer
  
  Dim W2InfoRec As W2FormType
  Dim W2InfoRecLen As Integer
  Dim NumInfoRec As Integer
  Dim W2InfoHandle As Integer
  
  Dim W2RSRec As W2ElectronicSubRS
  Dim W2RSLen As Integer
    
  Dim x As Integer, y As Integer
  Dim RptName$
  Dim RptHandle As Integer
  Dim TotNumOfRWS As Integer
  Dim TotWgsTips As Double
  Dim TotFedTax As Double
  Dim TotSSWages As Double
  Dim TotSSTax As Double
  Dim TotMedWages As Double
  Dim TotMedTax As Double
  Dim TotSSTips As Double
  Dim TotAdvEIC As Double
  Dim TotDepCare As Double
  Dim TotDefr401k As Double
  Dim TotDefr403b As Double
  Dim TotDefr408k6 As Double
  Dim TotDefr457b As Double
  Dim TotDefr501c18D As Double
  Dim TotNQPlan457 As Double
  Dim TotNQPNon457 As Double
  Dim TotLifIns As Double
  Dim TotThrdPrtySck As Double
  Dim TotNonStaStck As Double
  Dim TotAlloTips As Double
  Dim TotRoth401K As Double 'added for 2006
  Dim ThisLen As Integer
  Dim Blank5$, Blank3$
  Dim Blank9$, Blank11$
  Dim Blank12$, Blank1$
  Dim Blank56$, Blank23$, Blank113$
  Dim Blank176$, Blank15$, Blank143$
  Dim Blank30$, Blank158$, Blank45$
  Dim Blank496$, Blank291$, Blank240$
  Dim Blank22$, Blank57$, Blank128$
  Dim Blank6$, Blank2$, Blank336$
  Dim Blank4$, NumOfROs As Integer
  Dim ThisFile$, ThisYear$, ThisF$
  
  Dim Zip5 As String
  Dim Zip4 As String
  Dim StaWage As String
  Dim StaTax As String
  
  W2InfoRecLen = Len(W2InfoRec)
  
  On Error GoTo ERRORSTUFF
  ThisFile = "\W2Report\W2ES" + QPTrim$(fptxtTaxYear.Text) + ".txt"
  If DirExists(StartPath + "\W2Report") Then 'OK...here if the directory exists then there
  'must be a file inside otherwise the regular Exist would fail
    If Exist(StartPath + ThisFile) Then 'So...the directory with at least one file does exist...then
    'is the file in question one of those inside the directory
      If MsgBox("The file W2Report\W2ES" + QPTrim$(fptxtTaxYear.Text) + " already exists. If you continue the existing file will be overwritten. Do you wish to continue anyway?", vbYesNo) = vbNo Then
        Close
        vaTabPro1.ActiveTab = 0
        fptxtEmpEIN.SetFocus
        Exit Sub
      End If
      KillFile (StartPath + ThisFile)
    End If
  Else 'Ok...the directory might show up but if nothing is inside then
  'it really does not exist so it's OK at this point to remove the empty directory
  'and recreate a new one or if it never existed then go ahead and make one
    MkDir StartPath + "\W2Report"
  End If
  
  Blank5 = String(5, " ")
  Blank3 = String(3, " ")
  Blank9 = String(9, " ")
  Blank11 = String(11, " ")
  Blank12 = String(12, " ")
  Blank1 = String(1, " ")
  Blank56 = String(56, " ")
  Blank23 = String(23, " ")
  Blank176 = String(176, " ")
  Blank15 = String(15, " ")
  Blank30 = String(30, " ")
  Blank158 = String(158, " ")
  Blank496 = String(496, " ")
  Blank291 = String(291, " ")
  Blank22 = String(22, " ")
  Blank57 = String(57, " ")
  Blank6 = String(6, " ")
  Blank2 = String(2, " ")
  Blank4 = String(4, " ")
  Blank128 = String(128, " ")
  Blank240 = String(240, " ")
  Blank45 = String(45, " ")
  Blank143 = String(143, " ")
  Blank113 = String(113, " ")
  
  RptName$ = StartPath + ThisFile
  RptHandle = FreeFile
  Open RptName$ For Output As #RptHandle
  
  If Not Exist("PRDATA\W2ESUBRA.DAT") Then
    MsgBox "W2 Electronic File Build aborted because no 'W2SUBRA.DAT' file could be located."
    Close
    Exit Sub
  End If
  
  OpenW2ESubRA RAHandle
  Get RAHandle, 1, W2RARec
  
  '  rec ident Sub Emp Num     Emp Pers ID           ReSub          ReSub WFID      software code
  Print #RptHandle, "RA"; W2RARec.EINNum; W2RARec.PersIDNum; W2RARec.ResubID; W2RARec.ReSubWFID; W2RARec.SftwrCode;
  LSet W2RARec.CmpnyName = W2RARec.CmpnyName
  LSet W2RARec.LocAddr = W2RARec.LocAddr
  LSet W2RARec.DelAddr = W2RARec.DelAddr
  LSet W2RARec.City = W2RARec.City
  '                   Company Name    Location Address Delivery Address       City          State
  Print #RptHandle, W2RARec.CmpnyName; W2RARec.LocAddr; W2RARec.DelAddr; W2RARec.City; W2RARec.State;
  
  '                    Zip        Zip Extension  Blank Foreign State/Prov Foreign Postal Code
  Print #RptHandle, W2RARec.Zip; W2RARec.ZipExt; Blank5; Blank23; Blank15;
  
  LSet W2RARec.SubmttrName = W2RARec.SubmttrName
  LSet W2RARec.SubLocAddr = W2RARec.SubLocAddr
  LSet W2RARec.SubDelAddr = W2RARec.SubDelAddr
  LSet W2RARec.SubCity = W2RARec.SubCity
  '            Country Code Submtr Name          Sub Location        Sub Delivery         Sub City
  Print #RptHandle, Blank2; W2RARec.SubmttrName; W2RARec.SubLocAddr; W2RARec.SubDelAddr; W2RARec.SubCity;
  
  '                 Submitter's State    Sub Zip          Sub ZipExt     Blank Foreign State/Province
  Print #RptHandle, W2RARec.SubState; W2RARec.SubZip; W2RARec.SubZipExt; Blank5; Blank23;
    
  LSet W2RARec.ContactName = W2RARec.ContactName
  LSet W2RARec.CntctPhone = W2RARec.CntctPhone
  LSet W2RARec.CntPhnExt = W2RARec.CntPhnExt
  
  '             Foreign Post Code Cntry Cd   Contact Name   Contact Phone     Contact Phone Ext
  Print #RptHandle, Blank15; Blank2; W2RARec.ContactName; W2RARec.CntctPhone; W2RARec.CntPhnExt;
  
  LSet W2RARec.CntEMail = W2RARec.CntEMail
  LSet W2RARec.CntFAX = W2RARec.CntFAX
  '
  Print #RptHandle, Blank3; W2RARec.CntEMail; Blank3; W2RARec.CntFAX; W2RARec.CntMethod; W2RARec.PrepCode; Blank12
    
  If Not Exist("PRDATA\W2ESUBRE.DAT") Then
    MsgBox "W2 Electronic File Build aborted because no 'W2SUBRE.DAT' file could be located."
    Close
    Exit Sub
  End If
  
  OpenW2ESubRE REHandle
  Get REHandle, 1, W2RERec
  Close REHandle
  '     RecID                Tax Year Agent Indicator Code       Eplr/Agnt EIN ID   Agent For EIN   Term Bus Indicator  Est Num  Other EIN
  Print #RptHandle, "RE"; W2RERec.TaxYear; W2RERec.AgentCode; W2RERec.EmprAgntEIN; W2RERec.EINAgent; W2RERec.TermBusInd; Blank4; W2RERec.OthEIN;
  LSet W2RERec.EmprName = W2RERec.EmprName
  LSet W2RERec.EmprLocAddr = W2RERec.EmprLocAddr
  LSet W2RERec.EmprDelAddr = W2RERec.EmprDelAddr
  LSet W2RERec.EmprCity = W2RERec.EmprCity
  '                   Employer Name  Empr Location Address   Empr Delivery Add    Employer City
  Print #RptHandle, W2RERec.EmprName; W2RERec.EmprLocAddr; W2RERec.EmprDelAddr; W2RERec.EmprCity;
  '                  Employer State     Employer Zip    Employer Zip Ex    Blank  Foreign State/Province
  Print #RptHandle, W2RERec.EmprState; W2RERec.EmprZip; W2RERec.EmprZipX;
  'Dale this
  Print #RptHandle, W2RERec.EmprType; String(4, " "); Blank23;
  '            Foreign Post Cntry  Emp Cd Tax Cd Third Party Sick Code  Blank (40/291)
  Print #RptHandle, Blank15; Blank2; "R"; Blank1; W2RERec.ThrdSckPay; Blank291
  
  '-----------------------------------------------------
  W2InfoHandle = FreeFile
  Open W2InfoFile For Random Shared As W2InfoHandle Len = W2InfoRecLen
  
  OpenW2ESubRW RWHandle
  NumOfRWRecs = LOF(RWHandle) / Len(W2RWRec)
  For x = 1 To NumOfRWRecs
  Get RWHandle, x, W2RWRec
  Get W2InfoHandle, x, W2InfoRec
  
  If Len(QPTrim$(W2RWRec.WageTips)) = 0 And Len(QPTrim$(W2RWRec.FedTax)) = 0 And Len(QPTrim$(W2RWRec.SSWages)) = 0 Then
    If Len(QPTrim$(W2RWRec.SSTax)) = 0 And Len(QPTrim$(W2RWRec.MedWages)) = 0 And Len(QPTrim$(W2RWRec.MedTax)) = 0 Then
      If Len(QPTrim$(W2RWRec.SSTips)) = 0 And Len(QPTrim$(W2RWRec.AdvEIC)) = 0 Then
        If Len(QPTrim$(W2RWRec.DepCare)) = 0 And Len(QPTrim$(W2RWRec.NQPlan457)) = 0 And Len(QPTrim$(W2RWRec.Defr401k)) = 0 Then
          If Len(QPTrim$(W2RWRec.Defr403b)) = 0 And Len(QPTrim$(W2RWRec.Defr408k6)) = 0 And Len(QPTrim$(W2RWRec.Defr457b)) = 0 Then
            If Len(QPTrim$(W2RWRec.Defr501c18D)) = 0 Then
              GoTo DontPrintIt
            End If
          End If
        End If
      End If
    End If
  End If
  TotNumOfRWS = TotNumOfRWS + 1
  LSet W2RWRec.EmpFName = W2RWRec.EmpFName
  LSet W2RWRec.EmpMName = W2RWRec.EmpMName
  LSet W2RWRec.EmpLName = W2RWRec.EmpLName
  LSet W2RWRec.EmpSuffix = W2RWRec.EmpSuffix
  '                Rec ID  Social Sec #       First Name      Middle Name         Last Name    Jr. or II or Esq., etc
  Print #RptHandle, "RW"; W2RWRec.EmpSSN; W2RWRec.EmpFName; W2RWRec.EmpMName; W2RWRec.EmpLName; W2RWRec.EmpSuffix;
  LSet W2RWRec.EmpAdd2 = W2RWRec.EmpAdd2
  LSet W2RWRec.EmpAdd1 = W2RWRec.EmpAdd1
  LSet W2RWRec.EmpCity = W2RWRec.EmpCity
  '                 Location Addrss  Delivery Addrss   Employee City    Employee State    Employee Zip     Emp Zip Ext
  Print #RptHandle, W2RWRec.EmpAdd2; W2RWRec.EmpAdd1; W2RWRec.EmpCity; W2RWRec.EmpState; W2RWRec.EmpZip; W2RWRec.EmpZipX;
  '                 Blank Foreign State/Province Foreign Post Cd  Cntry
  Print #RptHandle, Blank5; Blank23; Blank15; Blank2;
  TotWgsTips = TotWgsTips + Val(W2RWRec.WageTips)
  Call ZeroFill(W2RWRec.WageTips, 11)
  TotFedTax = TotFedTax + Val(W2RWRec.FedTax)
  Call ZeroFill(W2RWRec.FedTax, 11)
  TotSSWages = TotSSWages + Val(W2RWRec.SSWages)
  Call ZeroFill(W2RWRec.SSWages, 11)
  TotSSTax = TotSSTax + Val(W2RWRec.SSTax)
  Call ZeroFill(W2RWRec.SSTax, 11)
  TotMedWages = TotMedWages + Val(W2RWRec.MedWages)
  Call ZeroFill(W2RWRec.MedWages, 11)
  TotMedTax = TotMedTax + Val(W2RWRec.MedTax)
  Call ZeroFill(W2RWRec.MedTax, 11)
  '                  Emp Fed Wages      Emp Fed Tax     Emp SS Wages     Emp SS Tax  Emp Medicare Wages  Emp Med Tax
  Print #RptHandle, W2RWRec.WageTips; W2RWRec.FedTax; W2RWRec.SSWages; W2RWRec.SSTax; W2RWRec.MedWages; W2RWRec.MedTax;
  
  TotSSTips = TotSSTips + Val(W2RWRec.SSTips)
  Call ZeroFill(W2RWRec.SSTips, 11)
  TotAdvEIC = TotAdvEIC + Val(W2RWRec.AdvEIC)
  Call ZeroFill(W2RWRec.AdvEIC, 11)
  TotDepCare = TotDepCare + Val(W2RWRec.DepCare)
  Call ZeroFill(W2RWRec.DepCare, 11)
  TotDefr401k = TotDefr401k + Val(W2RWRec.Defr401k)
  Call ZeroFill(W2RWRec.Defr401k, 11)
  TotDefr403b = TotDefr403b + Val(W2RWRec.Defr403b)
  Call ZeroFill(W2RWRec.Defr403b, 11)
  TotDefr408k6 = TotDefr408k6 + Val(W2RWRec.Defr408k6)
  Call ZeroFill(W2RWRec.Defr408k6, 11)
  '                Emp Soc Sec Tips  Emp Adv EIC    Emp Dependent Cr Emp Deferred 401k Emp Deferr 403b  Emp Deferred 408k6
  Print #RptHandle, W2RWRec.SSTips; W2RWRec.AdvEIC; W2RWRec.DepCare; W2RWRec.Defr401k; W2RWRec.Defr403b; W2RWRec.Defr408k6;

  TotDefr457b = TotDefr457b + Val(W2RWRec.Defr457b)
  Call ZeroFill(W2RWRec.Defr457b, 11)
  TotDefr501c18D = TotDefr501c18D + Val(W2RWRec.Defr501c18D)
  Call ZeroFill(W2RWRec.Defr501c18D, 11)
  TotNQPlan457 = TotNQPlan457 + Val(W2RWRec.NQPlan457)
  Call ZeroFill(W2RWRec.NQPlan457, 11)
  TotNQPNon457 = TotNQPNon457 + Val(W2RWRec.NQPNot457)
  Call ZeroFill(W2RWRec.NQPNot457, 11)
  '                Emp Deferred 457b  Emp Defred 501c18D  Blank 11  Emp NQP Dist/Contr Blank 11  Emp NQP Not 457
'  Print #RptHandle, W2RWRec.Defr457b; W2RWRec.Defr501c18D; Blank11; W2RWRec.NQPlan457; Blank11; W2RWRec.NQPNot457;
  '8/25/04...the 2004 version of accuwage required a field with 11 zeros instead of the blank field
  'that was what the 2003 version wanted...the zeros assume that there are no employer contributions
  'to health savings plans
  '                                                      Military Combat Pay 'new for 2005
  Print #RptHandle, W2RWRec.Defr457b; W2RWRec.Defr501c18D; "00000000000"; W2RWRec.NQPlan457; "00000000000"; W2RWRec.NQPNot457;

  TotLifIns = TotLifIns + Val(W2RWRec.LifeIns)
  Call ZeroFill(W2RWRec.LifeIns, 11)
  TotNonStaStck = TotNonStaStck + Val(W2RWRec.NonStaStcks)
  Call ZeroFill(W2RWRec.NonStaStcks, 11)
  TotRoth401K = TotRoth401K + Val(W2RWRec.Roth401K)
  Call ZeroFill(W2RWRec.Roth401K, 11) 'added for 2006
  '                 No Tax Combat Blank 11 Employer Life Ins Cost Income from NonSta Stock Opts '     Blank 56
  Print #RptHandle, "00000000000"; Blank11; W2RWRec.LifeIns; W2RWRec.NonStaStcks; ' "00000000000"; Blank45; 'NEW FOR 2005
  '                 New for 2006  Roth To 401K 2006  Roth to 403B 2006
  Print #RptHandle, "00000000000"; W2RWRec.Roth401K; "00000000000"; Blank23;
  '                Empy Statutory ID Blank Empy Retirement ID Empy 3rd Party Sick ID  Blank 23
  Print #RptHandle, W2RWRec.StatuEmp; Blank1; W2RWRec.RetPlan; W2RWRec.ThrdSckPay; Blank23
  
  TotThrdPrtySck = TotThrdPrtySck + Val(W2RWRec.ThrdSckAmt)
'  If x = 269 Then Stop
  If Val(W2RWRec.RONum) > 0 Then
    OpenW2ESubRO ROHandle
    Get ROHandle, Val(W2RWRec.RONum), W2RORec
    TotAlloTips = TotAlloTips + Val(W2RORec.AllocTips)
    '
    Print #RptHandle, "RO"; Blank9; W2RORec.AllocTips; W2RORec.TaxOnTips;
    '
    Print #RptHandle, W2RORec.MedSavings; W2RORec.RetAcct; W2RORec.AdoptionX;
    '
    Print #RptHandle, W2RORec.UnSSLife; W2RORec.UnMedLife; Blank176; Blank1; Blank9;
    '                  PRico Wages     PRico Comm    PRico Allow     PRico Tips
    Print #RptHandle, "00000000000"; "00000000000"; "00000000000"; "00000000000";
    '                  PRico TotWgs   PRico TaxWH   PRico Ret Fund
    Print #RptHandle, "00000000000"; "00000000000"; "00000000000"; Blank11;
    '                  Pacific TWgs   Pac Tax WH
    Print #RptHandle, "00000000000"; "00000000000"; Blank128
  End If
  
    W2RSRec.RecID = "RS"
    W2RSRec.StateCode = GetStateCode4W2(W2RWRec.EmpState)
    W2RSRec.Fill1 = ""
    W2RSRec.SSN = W2RWRec.EmpSSN
    W2RSRec.FName = W2RWRec.EmpFName
    W2RSRec.MName = W2RWRec.EmpMName
    W2RSRec.LName = W2RWRec.EmpLName
    W2RSRec.Fill2 = ""
    W2RSRec.EmpAddr = W2RWRec.EmpAdd1
    W2RSRec.DelAddr = W2RWRec.EmpAdd1
    W2RSRec.City = W2RWRec.EmpCity
    W2RSRec.State = W2RWRec.EmpState
    Call MakeZipCode4RS(W2RWRec.EmpZip, Zip5, Zip4)
    W2RSRec.Zip5 = Zip5
    W2RSRec.ZipPlus4 = Zip4
    W2RSRec.Fill3 = ""
    W2RSRec.EmpAcctNo = ""
    W2RSRec.Fill4 = ""
    
    MakeStateWageTax W2InfoRec.StaWage, W2InfoRec.STATAXWH, StaWage, StaTax
    
    W2RSRec.StateWage = StaWage
    W2RSRec.StateTax = StaTax
    
    W2RSRec.Vested = ""
    W2RSRec.PadFill = ""
    
'*************
  
  Print #RptHandle, W2RSRec.RecID;
  Print #RptHandle, W2RSRec.StateCode;
    Print #RptHandle, W2RSRec.Fill1;
  Print #RptHandle, W2RSRec.SSN;
  Print #RptHandle, W2RSRec.FName;
  Print #RptHandle, W2RSRec.MName;
  Print #RptHandle, W2RSRec.LName;
    Print #RptHandle, W2RSRec.Fill2;
  Print #RptHandle, W2RSRec.EmpAddr;
  Print #RptHandle, W2RSRec.DelAddr;
  Print #RptHandle, W2RSRec.City;
  Print #RptHandle, W2RSRec.State;
  Print #RptHandle, W2RSRec.Zip5;
  Print #RptHandle, W2RSRec.ZipPlus4;
    Print #RptHandle, W2RSRec.Fill3;
  Print #RptHandle, W2RSRec.EmpAcctNo;
    Print #RptHandle, W2RSRec.Fill4;
  Print #RptHandle, W2RSRec.StateWage;
  Print #RptHandle, W2RSRec.StateTax;
  Print #RptHandle, W2RSRec.Vested;
  Print #RptHandle, W2RSRec.PadFill

  'Close ROHandle

DontPrintIt:
  Next x
'--------------------------------------------------------
  If TotNumOfRWS = 0 Then
    MsgBox "There are no W2 employee records processed at this time. Please make sure you have extracted W2 data for valid employees for the current year. W2 Electronic File processing aborted."
    Close
    KillFile ("PRDATA\W2ESUBRW.DAT")
    KillFile ("PRDATA\W2ESUBRO.DAT")
    Exit Sub
  End If
  
  W2RTRec.NumOfRWS = CStr(TotNumOfRWS)
  Call ZeroFill(W2RTRec.NumOfRWS, 7)
  W2RTRec.WagesTips = CStr(TotWgsTips)
  Call ZeroFill(W2RTRec.WagesTips, 15)
  W2RTRec.FedTax = CStr(TotFedTax)
  Call ZeroFill(W2RTRec.FedTax, 15)
  W2RTRec.SocWages = CStr(TotSSWages)
  Call ZeroFill(W2RTRec.SocWages, 15)
  W2RTRec.SocTax = CStr(TotSSTax)
  Call ZeroFill(W2RTRec.SocTax, 15)
  W2RTRec.MedWages = CStr(TotMedWages)
  Call ZeroFill(W2RTRec.MedWages, 15)
  W2RTRec.MedTax = CStr(TotMedTax)
  Call ZeroFill(W2RTRec.MedTax, 15)
  W2RTRec.SocTips = CStr(TotSSTips)
  Call ZeroFill(W2RTRec.SocTips, 15)
  W2RTRec.AdvEIC = CStr(TotAdvEIC)
  Call ZeroFill(W2RTRec.AdvEIC, 15)
  W2RTRec.DepCare = CStr(TotDepCare)
  Call ZeroFill(W2RTRec.DepCare, 15)
  W2RTRec.Defr401k = CStr(TotDefr401k)
  Call ZeroFill(W2RTRec.Defr401k, 15)
  W2RTRec.Defr403b = CStr(TotDefr403b)
  Call ZeroFill(W2RTRec.Defr403b, 15)
  W2RTRec.Defr408k6 = CStr(TotDefr408k6)
  Call ZeroFill(W2RTRec.Defr408k6, 15)
  W2RTRec.Defr457b = CStr(TotDefr457b)
  Call ZeroFill(W2RTRec.Defr457b, 15)
  W2RTRec.Defr501c18D = CStr(TotDefr501c18D)
  Call ZeroFill(W2RTRec.Defr501c18D, 15)
  W2RTRec.NQPlan457 = CStr(TotNQPlan457)
  Call ZeroFill(W2RTRec.NQPlan457, 15)
  W2RTRec.NQPNot457 = CStr(TotNQPNon457)
  Call ZeroFill(W2RTRec.NQPNot457, 15)
  W2RTRec.GrpTerm = CStr(TotLifIns)
  Call ZeroFill(W2RTRec.GrpTerm, 15)
  W2RTRec.ThrdTaxPay = CStr(TotThrdPrtySck)
  Call ZeroFill(W2RTRec.ThrdTaxPay, 15)
  W2RTRec.NonStatStk = CStr(TotNonStaStck)
  Call ZeroFill(W2RTRec.NonStatStk, 15)
  W2RTRec.Roth401K = CStr(TotRoth401K)
  Call ZeroFill(W2RTRec.Roth401K, 15)
  
  '                Rec ID    emp Rec Cnt        Total Wages      Total Fed Tax
  Print #RptHandle, "RT"; W2RTRec.NumOfRWS; W2RTRec.WagesTips; W2RTRec.FedTax;
  '                 Tot Social Wages  Total Soc Tax   Total Medicare Wgs Total Med Tax
  Print #RptHandle, W2RTRec.SocWages; W2RTRec.SocTax; W2RTRec.MedWages; W2RTRec.MedTax;
  '                 Tot Social Tips   Total AdvEic Total Dependant Cr  Total Deferred 401k
  Print #RptHandle, W2RTRec.SocTips; W2RTRec.AdvEIC; W2RTRec.DepCare; W2RTRec.Defr401k;
  '                Tot Deferred 403b Tot Deferred 408k6 Tot Deferred 457b Tot Deferred 501c18D
  Print #RptHandle, W2RTRec.Defr403b; W2RTRec.Defr408k6; W2RTRec.Defr457b; W2RTRec.Defr501c18D;
  '8/25/04...the 2004 version of accuwage required a field with 15 zeros instead of the blank field
  'that was what the 2003 version wanted...the zeros assume that there are no employer contributions
  'to health savings plans
  
  '              Military Combat Total new for 2005
  Print #RptHandle, "000000000000000"; W2RTRec.NQPlan457; "000000000000000"; W2RTRec.NQPNot457;
  '         No Tax Combat new for 2005 Blank15 Total Life Ins   Total 3rd Sick Pay Total NonStatutory Stock Options
  Print #RptHandle, "000000000000000"; Blank15; W2RTRec.GrpTerm; W2RTRec.ThrdTaxPay; W2RTRec.NonStatStk;
  '         Defer Sec409A new for 2005
  Print #RptHandle, "000000000000000"; 'Blank143
  ' New for 2006:    Total Roth 401K      Total Roth 403B
  Print #RptHandle, W2RTRec.Roth401K; "000000000000000"; Blank113
  
  OpenW2ESubRT RTHandle
  Put RTHandle, 1, W2RTRec
  Close RTHandle
  '-------------------------------------------------
  If Exist("PRDATA\W2ESUBRO.DAT") Then
    OpenW2ESubRO ROHandle
    NumOfROs = LOF(ROHandle) / Len(W2RORec)
    Close ROHandle
    If NumOfROs > 0 Then
      W2RURec.NumOfROs = CStr(NumOfROs)
      Call ZeroFill(W2RURec.NumOfROs, 7)
      W2RURec.AllocTips = TotAlloTips
      Call ZeroFill(W2RURec.AllocTips, 15)
      W2RURec.TaxOnTips = "000000000000000"
      W2RURec.MedSavings = "000000000000000"
      W2RURec.RetAcct = "000000000000000"
      W2RURec.AdoptionX = "000000000000000"
      W2RURec.UnSSLife = "000000000000000"
      W2RURec.UnMedLife = "000000000000000"
      OpenW2ESubRU RUHandle
      Put RUHandle, 1, W2RURec
      Close RUHandle
      '               Rec ID   Total ROs       Total Allocated Tips Uncollected Tips  Total Medical Sav
      Print #RptHandle, "RU"; W2RURec.NumOfROs; W2RURec.AllocTips; W2RURec.TaxOnTips; W2RURec.MedSavings;
      '                Simple Retirement Tot Adoption Exps Uncollected SS Ins Uncollected Med Ins
      Print #RptHandle, W2RURec.RetAcct; W2RURec.AdoptionX; W2RURec.UnSSLife; W2RURec.UnMedLife; Blank240;
      '                    PRico Wages     PRico Commissions     PRico Allow         PRico Tips      PRico Totals
      Print #RptHandle, "000000000000000"; "000000000000000"; "000000000000000"; "000000000000000"; "000000000000000";
      '                   PRico Tax WH      PRico Ret Fund     Pacific Totals      Pacific Tax WH
      Print #RptHandle, "000000000000000"; "000000000000000"; "000000000000000"; "000000000000000"; Blank23
    End If
  End If
  '-------------------------------------------------
  W2RFRec.NumOfRWS = CStr(TotNumOfRWS)
  Call ZeroFill(W2RFRec.NumOfRWS, 9)
  '                Rec ID  Blank   Number of RW Records
  Print #RptHandle, "RF"; Blank5; W2RFRec.NumOfRWS;
  '
  Print #RptHandle, Blank496
  
  
  OpenW2ESubRF RFHandle
  Put RFHandle, 1, W2RFRec
  Close RptHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmW2ElecSub", "BuildEFile", Erl)
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
    MsgBox "ERROR: W2 submission file building aborted."
End Sub

Function GetStateCode4W2$(StateAbbr$)
    Dim StaCod As String

   Select Case StateAbbr$
    Case "AL"
    StaCod = "01"
    Case "AK"
    StaCod = "02"
    Case "AZ"
    StaCod = "04"
    Case "AR"
    StaCod = "05"
    Case "CA"
    StaCod = "06"
    Case "CO"
    StaCod = "08"
    Case "CT"
    StaCod = "09"
    Case "DE"
    StaCod = "10"
    Case "DC"
    StaCod = "11"
    Case "FL"
    StaCod = "12"
    Case "GA"
    StaCod = "13"
    Case "HI"
    StaCod = "15"
    Case "ID"
    StaCod = "16"
    Case "IL"
    StaCod = "17"
    Case "IN"
    StaCod = "18"
    Case "IA"
    StaCod = "19"
    Case "KS"
    StaCod = "20"
    Case "KY"
    StaCod = "21"
    Case "LA"
    StaCod = "22"
    Case "ME"
    StaCod = "23"
    Case "MD"
    StaCod = "24"
    Case "MA"
    StaCod = "25"
    Case "MI"
    StaCod = "26"
    Case "MN"
    StaCod = "27"
    Case "MS"
    StaCod = "28"
    Case "MO"
    StaCod = "29"
    Case "MT"
    StaCod = "30"
    Case "NE"
    StaCod = "31"
    Case "NV"
    StaCod = "32"
    Case "NH"
    StaCod = "33"
    Case "NJ"
    StaCod = "34"
    Case "NM"
    StaCod = "35"
    Case "NY"
    StaCod = "36"
    Case "NC"
    StaCod = "37"
    Case "ND"
    StaCod = "38"
    Case "OH"
    StaCod = "39"
    Case "OK"
    StaCod = "40"
    Case "OR"
    StaCod = "41"
    Case "PA"
    StaCod = "42"
    Case "RI"
    StaCod = "44"
    Case "SC"
    StaCod = "45"
    Case "SD"
    StaCod = "46"
    Case "TN"
    StaCod = "47"
    Case "TX"
    StaCod = "48"
    Case "UT"
    StaCod = "49"
    Case "VT"
    StaCod = "50"
    Case "VA"
    StaCod = "51"
    Case "WA"
    StaCod = "53"
    Case "WV"
    StaCod = "54"
    Case "WI"
    StaCod = "55"
    Case "WY"
    StaCod = "56"
   End Select
   GetStateCode4W2$ = StaCod

End Function

Sub MakeZipCode4RS(EmpZip, Zip5, Zip4)
    Dim TempZip As String
    Dim ZipLen As Integer
    Dim DashPos As Integer
    TempZip = EmpZip
    TempZip = QPTrim$(TempZip)
    
    ZipLen = Len(TempZip)
    If ZipLen > 6 Then
        DashPos = InStr(TempZip, "-")
        If DashPos > 0 Then
          TempZip = Left$(TempZip, DashPos - 1) + Mid$(TempZip$, DashPos + 1)
        End If
    ElseIf ZipLen = 6 Then
        TempZip = Left$(TempZip, 5)
    End If
    
    ZipLen = Len(TempZip)
    
    Select Case ZipLen
    Case 5 To 8
      Zip5 = Left$(TempZip, 5)
      Zip4 = "    "
    Case 9
      Zip5 = Left$(TempZip, 5)
      Zip4 = Right$(TempZip, 4)
    Case Else
      Zip5 = "     "
      Zip4 = "    "
    End Select
End Sub

Sub MakeStateWageTax(StateWage As Double, StateTax As Double, StaWage As String, StaTax As String)
  Dim PerPos As Integer
  Dim Bucks As String
  Dim Cents As String
  Dim Make11 As String
  
  StaWage = CStr(StateWage)
  StaTax = CStr(StateTax)
  
  StaWage = QPTrim$(StaWage)
  StaTax = QPTrim$(StaTax)
  
  PerPos = InStr(StaWage, ".")
  If PerPos > 0 Then
    Bucks = Left$(StaWage, PerPos - 1)
    Cents = Mid$(StaWage, PerPos + 1)
    If Len(Cents) = 1 Then
      Cents = Cents + "0"
    End If
  Else
    Bucks = StaWage
    Cents = "00"
  End If
  Make11 = "00000000000" + Bucks + Cents
  StaWage = Right$(Make11, 11)
  
  PerPos = InStr(StaTax, ".")
  If PerPos > 0 Then
    Bucks = Left$(StaTax, PerPos - 1)
    Cents = Mid$(StaTax, PerPos + 1)
    If Len(Cents) = 1 Then
      Cents = Cents + "0"
    End If
  Else
    Bucks = StaTax
    Cents = "00"
  End If
  Make11 = "00000000000" + Bucks + Cents
  StaTax = Right$(Make11, 11)
  
End Sub
