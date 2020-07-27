VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxMasterBalList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Master Balance Listing"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxMasterBalList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6276
      Left            =   1908
      TabIndex        =   5
      Top             =   1242
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11070
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxMasterBalList.frx":08CA
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   3525
         TabIndex        =   2
         Top             =   2730
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
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
         Columns         =   0
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
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
         ColDesigner     =   "frmTaxMasterBalList.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   3
         Top             =   3285
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         Columns         =   0
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
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
         ColDesigner     =   "frmTaxMasterBalList.frx":0BE1
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   4
         Top             =   3870
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         BackColor       =   16777215
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
         ColumnEdit      =   -1
         ColumnBound     =   -1
         Style           =   2
         MaxDrop         =   8
         ListWidth       =   -1
         EditHeight      =   -1
         GrayAreaColor   =   -2147483633
         ListLeftOffset  =   0
         ComboGap        =   -2
         MaxEditLen      =   5
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
         ColDesigner     =   "frmTaxMasterBalList.frx":0EDC
      End
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   405
         Left            =   3525
         TabIndex        =   1
         Top             =   2160
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
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
         BackColor       =   16777215
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
         ColumnEdit      =   -1
         ColumnBound     =   -1
         Style           =   2
         MaxDrop         =   8
         ListWidth       =   -1
         EditHeight      =   -1
         GrayAreaColor   =   -2147483633
         ListLeftOffset  =   0
         ComboGap        =   -2
         MaxEditLen      =   5
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
         ColDesigner     =   "frmTaxMasterBalList.frx":11D7
      End
      Begin LpLib.fpCombo fpcmbTaxYear 
         Height          =   405
         Left            =   3720
         TabIndex        =   0
         Top             =   1560
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
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
         Columns         =   0
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
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
         ColDesigner     =   "frmTaxMasterBalList.frx":14D2
      End
      Begin VB.CheckBox chkZeroBal 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Check to Include Zero Balances"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2160
         TabIndex        =   14
         Top             =   4440
         Width           =   3732
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5250
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxMasterBalList.frx":17CD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4275
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmTaxMasterBalList.frx":19AB
         Top             =   5250
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxMasterBalList.frx":1A56
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Account Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1395
         TabIndex        =   13
         Top             =   2835
         Width           =   1950
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Year:"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1275
         TabIndex        =   11
         Top             =   3390
         Width           =   1500
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1530
         Top             =   315
         Width           =   4905
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Master Balance Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1800
         TabIndex        =   10
         Top             =   450
         Width           =   4335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
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
         Left            =   1470
         TabIndex        =   9
         Top             =   3960
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3780
         Left            =   1008
         Top             =   1248
         Width           =   5976
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detail/Summary:"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2280
         Width           =   1905
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6540
      Left            =   1788
      Top             =   1098
      Width           =   8052
   End
End
Attribute VB_Name = "frmTaxMasterBalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Town$
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim BeenPrinted As Boolean
  Dim TotCustCount As Integer
  Dim LngLastCust As Long
  
Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  TotCustCount = 0
  If fpcmbDetSum.Text = "Summary" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphics
    Else
      frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      Call PrintText
    End If
  Else
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsDet
    Else
      frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      Call PrintTextDet
    End If
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
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpMasterBalance
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxMasterBalList.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim YrCnt As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
'  Dim SaveDate$
'  'on error goto ERRORSTUFF
  
  UseOpt = "N"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.Name)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  frmTaxLoadReport.Label1.Caption = "Loading Years"
  frmTaxLoadReport.Show
  DoEvents
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    'Calabash fix 7/18/07
'    If TaxTrans.TaxYear = 8224 Then
'      SaveDate = MakeRegDate(TaxTrans.TransDate)
'      TaxTrans.TaxYear = 2002
'      Put TTHandle, x, TaxTrans
'    End If
    If YrCnt = 0 Then
      If TaxTrans.TaxYear > 0 Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    End If
  Next x
  Close TTHandle
  
  BigYr = 0
  For x = 1 To YrCnt
    If Years(x) > BigYr Then
      BigYr = Years(x)
    End If
  Next x
  
  Nextx = 1
  ThisBigYr = BigYr + 1
  Do While Nextx <= YrCnt
    For x = Nextx To YrCnt
      If Years(x) < ThisBigYr Then
        ThisBigYr = Years(x)
        Thisx = x
      End If
    Next x
    HoldYr = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldYr
    Nextx = Nextx + 1
    ThisBigYr = BigYr + 1
  Loop
    
  fpcmbTaxYear.Text = "All"
  fpcmbTaxYear.AddItem "All"
  
  For x = YrCnt To 1 Step -1
    fpcmbTaxYear.AddItem CStr(Years(x))
  Next x
  
  Unload frmTaxLoadReport
  fpcmbIncInactive.Text = "Both"
  fpcmbIncInactive.AddItem "Both"
  fpcmbIncInactive.AddItem "Active Only"
  fpcmbIncInactive.AddItem "Inactive Only"
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  fpcmbDetSum.Text = "Summary"
  fpcmbDetSum.AddItem "Detail"
  fpcmbDetSum.AddItem "Summary"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMasterBalList", "LoadMe", Erl)
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

Private Sub fpcmbDetSum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDetSum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDetSum.ListIndex = -1
  End If
  If fpcmbDetSum.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbIncInactive.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_Change()
  If ThisOpt <> "" Then
    If InStr(fpcmbPrintOrder.Text, ThisOpt) Then
      UseOpt = "Y"
    Else
      UseOpt = "N"
    End If
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTaxYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxYear.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxYear.ListIndex = -1
  End If
  If fpcmbTaxYear.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDetSum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim CustTotBal As Double
  Dim TestBal As Double
  Dim One As Integer
  Dim AHandle As Integer
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  TotCustCount = 0
  LngLastCust = 0
  
'  One = 1
'  AHandle = FreeFile
'  Open "masterbal.dat" For Output As AHandle
'  Print #AHandle, One
'  Close AHandle
'  'on error goto ERRORSTUFF
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  
  dlm$ = "~"
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If

  RptFile$ = "TAXRPTS\TXMSTBAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
'    If TaxCust.Acct = 661 Then Stop
    Balance = 0 'added 7/19/06
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    
    'If TaxCust.Acct = 1032 Then Stop
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
      CustTotBal = TestBal
    Else
      GoSub GetYearBalance
      CustTotBal = TestBal
    End If
    
    
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      ThisRec = TaxCust.LastTrans
      If fpcmbTaxYear.Text = "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      ElseIf fpcmbTaxYear.Text <> "All" Then
        OP = CustTotBal ' GetCustBalanceForYear(CustRec, CInt(fpcmbTaxYear.Text))
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      End If
      Dim SaveYear As Integer
      Dim TaxDate$
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        'Calabash fix '7/18/06 data below was changed to accommodate two type
        'of problems...1) transactions that belonged to bills with good tax years
        ' and 2) transactions that belonged to bills with tax year 8224...went to
        'loadme and fixed remainig transactions with 8224 tax year saved
'        If TaxTrans.TaxYear = 8224 Then Stop 'And TaxTrans.TranType = 1 Then
'          Get TTHandle, TaxTrans.BelongTo, TaxTrans
'          TaxDate = MakeRegDate(TaxTrans.TransDate)
'          SaveYear = TaxTrans.TaxYear
''          Debug.Print CStr(ThisRec) + "  " + CStr(TaxTrans.TaxYear) + " Billing Rec = " + CStr(TaxTrans.BelongTo)
''          Stop
'          Get TTHandle, ThisRec, TaxTrans
'          TaxTrans.TaxYear = 2002
'          Put TTHandle, ThisRec, TaxTrans
'        End If
        If OP < 0 Then
          Balance = OP
          TaxTrans.LastTrans = 0
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text) 'tax year balance will be the balance
            'for the selected year
          End If
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
'          TaxTrans.CustomerRec = TaxTrans.CustomerRec
'          TaxTrans.BelongTo = TaxTrans.BelongTo
          
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero    'Jumps into an If conditional branch....horror!
          End If
          If Balance <> 0 Then
GoWithZero: 'BOB
            GBal = OldRound(GBal + Balance#)
            If YrCnt = 0 Then '
              YrCnt = YrCnt + 1
              ReDim Preserve Years(1 To YrCnt) As Integer
              Years(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve YearBal(1 To YrCnt) As Double
              YearBal(YrCnt) = Balance#
            Else
              For y = 1 To YrCnt
                If Years(y) = TaxTrans.TaxYear Then '
                  YearBal(y) = OldRound(YearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > YrCnt Then '
                YrCnt = YrCnt + 1
                ReDim Preserve Years(1 To YrCnt) As Integer
                Years(YrCnt) = TaxTrans.TaxYear
                ReDim Preserve YearBal(1 To YrCnt) As Double
                YearBal(YrCnt) = Balance#
              End If '
            End If '
'g total for year
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(YrCnt) = Balance#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
              End If '
            End If '
            
          BigYr = 0
          Nextz = 1
          For z = 1 To YrCnt
            If Years(z) > BigYr Then
              BigYr = Years(z)
            End If
          Next z
          
          ThisBigYr = BigYr + 1
          Nextz = 1
          Do While Nextz <= YrCnt
            For z = Nextz To YrCnt
              If Years(z) < ThisBigYr Then
                ThisBigYr = Years(z)
                Thisz = z
              End If
            Next z
            HoldYr = Years(Nextz)
            HoldBal = YearBal(Nextz)
            Years(Nextz) = Years(Thisz)
            YearBal(Nextz) = YearBal(Thisz)
            Years(Thisz) = HoldYr
            YearBal(Thisz) = HoldBal
            Nextz = Nextz + 1
            ThisBigYr = BigYr + 1
          Loop
        End If
      End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
      For z = YrCnt To 1 Step -1
'        If Years(z) = 2006 Then
'          Print #AHandle, CStr(CustRec) + "~" + CustName + "~" + Using$("#####.00", YearBal(z))
'        End If
        
        If LngLastCust <> CustRec Then
          LngLastCust = CustRec
          TotCustCount = TotCustCount + 1
        End If

        TransCnt = TransCnt + 1
        '                   0            1             2            3               4               5             6
        Print #RptHandle, Town; dlm; CustName; dlm; CustRec; dlm; Years(z); dlm; YearBal(z); dlm; TotCustCount; dlm; GBal; dlm;
        If UseOpt = "Y" Then
          '                    7                     8                           9               10               11               12
          Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm; CustTotBal; dlm; OverPay; dlm; fpcmbTaxYear.Text
        Else
          '                  7        8           9                10              11               12
          Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; CustTotBal; dlm; OverPay; dlm; fpcmbTaxYear.Text
        End If
      Next z
    End If
Inactive:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  'MsgBox (CStr(TotCustCount))
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  Close
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  SubRptFile$ = "TAXRPTS\TXMSTBALSUB.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
 
  If InStr(CStr(OverPay), "E") Then OverPay = 0
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      '                        0               1             2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; 1
    Else
      '                        0               1             2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; 0
    End If
  Next x
  
  Close
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
'  Close AHandle
        
  arTaxMasterBalSum.Show
  
  Exit Sub
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  '011713 Changed to correct issue with credit balances...
  If TestBal = 0 And OverPaid > 0 Then
    TestBal = -OverPaid
  Else
    TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  End If
  
  'TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
    
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  Return

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMasterBalList", "PrintGraphics", Erl)
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

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim FF$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim ThisAcct$
  Dim ThatAcct$
  Dim CustTotBal As Double
  Dim TestBal As Double
  Dim PrintCHeader As Boolean
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  'Dim TotCustCount As Integer
  TotCustCount = 1
  
'  'on error goto ERRORSTUFF
  PrintCHeader = False
  FF$ = Chr(12)
  MaxLines = 58
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
'  Else
'    MsgBox ("using account")
  
  End If

  RptFile$ = "TAXRPTS\TXMSTBAL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  GoSub PrintHeader
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0 'added 7/19/06
    PrintCHeader = True
    If TaxCust.Deleted <> 0 Then GoTo Inactive 'SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    
    ThisAcct = CStr(TaxCust.Acct)
    TestBal = 0
    If fpcmbTaxYear.Text <> "All" Then
      If TaxCust.LastTrans > 0 Then 'GoSub GetYearBalance
        GoSub GetYearBalance
        CustTotBal = TestBal 'GetCustBalanceForYear(CustRec, fpcmbTaxYear.Text)
      End If
      If CustTotBal = 0 Then
        If chkZeroBal.Value = 0 Then
          GoTo Inactive
        End If
      End If
    Else
      GoSub GetAllBalance
      CustTotBal = TestBal 'GetCustBalance(CustRec, -1)
      If CustTotBal = 0 Then
        If chkZeroBal.Value = 0 Then
          GoTo Inactive
        End If
      End If
    End If
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
'      ThisRec = TaxCust.LastTrans
      If fpcmbTaxYear.Text = "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      ElseIf fpcmbTaxYear.Text <> "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      End If
      If ThatAcct <> ThisAcct Then
        If LineCnt > MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          If LineCnt <> 6 Then
            Print #RptHandle,
            LineCnt = LineCnt + 1
          End If
        End If
        If TestBal <> 0 Or CustTotBal <> 0 Or chkZeroBal.Value = 1 And PrintCHeader = True Then
          GoSub PrintCustHeader
          PrintCHeader = False
        End If
        ThatAcct = ThisAcct
      End If
      
      ThisRec = TaxCust.LastTrans
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If OP < 0 Then
          TaxTrans.LastTrans = 0
          Balance = OP
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text) 'tax year balance will be the balance
            'for the selected year
          End If
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If YrCnt = 0 Then '
              YrCnt = YrCnt + 1
              ReDim Preserve Years(1 To YrCnt) As Integer
              Years(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve YearBal(1 To YrCnt) As Double
              YearBal(YrCnt) = Balance#
            Else
              For y = 1 To YrCnt
                If Years(y) = TaxTrans.TaxYear Then '
                  YearBal(y) = OldRound(YearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > YrCnt Then '
                YrCnt = YrCnt + 1
                ReDim Preserve Years(1 To YrCnt) As Integer
                Years(YrCnt) = TaxTrans.TaxYear
                ReDim Preserve YearBal(1 To YrCnt) As Double
                YearBal(YrCnt) = Balance#
              End If '
            End If '
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(YrCnt) = Balance#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
              End If '
            End If '
            
          BigYr = 0
          Nextz = 1
          For z = 1 To YrCnt
            If Years(z) > BigYr Then
              BigYr = Years(z)
            End If
          Next z
          
          ThisBigYr = BigYr + 1
          Nextz = 1
          Do While Nextz <= YrCnt
            For z = Nextz To YrCnt
              If Years(z) < ThisBigYr Then
                ThisBigYr = Years(z)
                Thisz = z
              End If
            Next z
            HoldYr = Years(Nextz)
            HoldBal = YearBal(Nextz)
            Years(Nextz) = Years(Thisz)
            YearBal(Nextz) = YearBal(Thisz)
            Years(Thisz) = HoldYr
            YearBal(Thisz) = HoldBal
            Nextz = Nextz + 1
            ThisBigYr = BigYr + 1
          Loop
        End If
      End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      
      Loop
            'TotCustCount = TotCustCount + 1
      For z = YrCnt To 1 Step -1
        TransCnt = TransCnt + 1
        Print #RptHandle, Tab(8); "Tax Year: " + Using$("###0", Years(z)); Tab(57); "Year Total: " + Using$("$###,##0.00", YearBal(z))
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If UseOpt = "Y" Then
          Print #RptHandle, Tab(10); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
        End If
      Next z
      If YrCnt > 0 Then
        Print #RptHandle, String(79, "-")
        LineCnt = LineCnt + 1 '2/15/06
      End If
    End If
Inactive:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  Print #RptHandle, FF$
  
  If InStr(CStr(OverPay), "E") Then OverPay = 0
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Summary"
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  Print #RptHandle, String(79, "-")
  
  Print #RptHandle, "Total Entries: " + Using$("##,##0", TotCustCount)
  Print #RptHandle, "Total Tax Balance:  " + Using$("$###,###,##0.00", GBal)
  If fpcmbTaxYear.Text = "All" Then
    Print #RptHandle, "Total Over Payment: " + Using$("$###,###,##0.00", OverPay)
  Else
    Print #RptHandle,
  End If
  Print #RptHandle,
  Print #RptHandle, "Tax Totals By Year"
  Print #RptHandle, Tab(12); "Tax Year"; Tab(26); "Amount Owed" '; Tab(42); "Over Payment"
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      Print #RptHandle, Tab(15); Using$("###0", GYears(x)); Tab(24); Using$("$#,###,##0.00", GYearBal(x)) '; Tab(44); Using$("$##,##0.00", OverPay)
    Else
      Print #RptHandle, Tab(15); Using$("###0", GYears(x)); Tab(24); Using$("$#,###,##0.00", GYearBal(x))
    End If
  Next x
  Print #RptHandle, FF$
  Close
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
        
'  MsgBox (CStr(TotCustCount))
  ViewPrint RptFile, "Tax Master Balance Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Summary"
  Print #RptHandle, Town
  If ActiveFlag = "B" Then
    Print #RptHandle, "Customer Status: " + "Active And Inactive"; Tab(65); "(I) = Inactive"
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  End If
  Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name" '; Tab(59); "Tax Year"; Tab(73); "Balance"
  Print #RptHandle, String(79, "-")
  LineCnt = 6
  
  Return
  
PrintCustHeader:
  Print #RptHandle, CustRec; Tab(8); CustName; Tab(48); "Total Balance: "; Tab(65); Using$("$###,###,##0.00", CustTotBal)
  Print #RptHandle, String(79, ".")
  LineCnt = LineCnt + 2
  TotCustCount = TotCustCount + 1
  'stophere
  Return

GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMasterBalList", "PrintText", Erl)
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

Private Sub PrintGraphicsDet()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPrincBal As Double
  Dim ThisIntBal As Double
  Dim ThisAdvBal As Double
  Dim ThisLateListBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim PinCnt As Integer
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim TestBal#
  Dim ThisOP$
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '9/17/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  Dim Dif As Double '9/19/07
  
'  'on error goto ERRORSTUFF
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  
  dlm$ = "~"
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If

  RptFile$ = "TAXRPTS\TXMSTBALDET.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPrincBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GAdvBal(1 To 1) As Double
  ReDim GLateListBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
'      If CustRec = 3362 Then Stop
    End If
   
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If

    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    CustTotBal = TestBal#
    OP = TestBal#
    
    If InStr(CStr(CustTotBal), "E") Then CustTotBal = 0
    If TaxCust.LastTrans > 0 Then
      YrCnt = 0
      PinCnt = 0
      ReDim YearBal(1 To 1) As Double
      ReDim Years(1 To 1) As Integer
      ReDim PrincBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim AdvBal(1 To 1) As Double
      ReDim LateListBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        Balance = 0 'added 7/19/06
'If TaxTrans.TaxYear = 2006 Then Stop
        ThisPrincBal = 0 'added 1/17/07
        ThisIntBal = 0 'added 1/17/07
        ThisAdvBal = 0 'added 1/17/07
        ThisLateListBal = 0 'added 1/17/07
        ThisOpt1Bal = 0 'added 1/17/07
        ThisOpt2Bal = 0 'added 1/17/07
        ThisOpt3Bal = 0 'added 1/17/07
        If CustTotBal < 0 Then
          Balance# = CustTotBal
          ThisPrincBal = CustTotBal 'added 1/17/07
          TaxTrans.LastTrans = 0
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          ThisPrincBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          ThisPrincBal = OldRound(ThisPrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisAdvBal = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          ThisLateListBal = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)) ' changed on 1/17/07 + TaxTrans.DiscAmt))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.RealPin) = "-1" Then
              ThisPin = "-1"
              PropType = "MOCK"
            ElseIf QPTrim$(TaxTrans.RealPin) <> "0" And QPTrim$(TaxTrans.RealPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.RealPin)
              PropType = "REAL"
            ElseIf QPTrim$(TaxTrans.PersPin) <> "0" And QPTrim$(TaxTrans.PersPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.PersPin)
              PropType = "PERSONAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPrincBal(1 To GYrCnt) As Double
              GPrincBal(GYrCnt) = ThisPrincBal#
              GPrincTot = ThisPrincBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              ReDim Preserve GAdvBal(1 To GYrCnt) As Double
              GAdvBal(GYrCnt) = ThisAdvBal#
              GAdvTot = ThisAdvBal#
              ReDim Preserve GLateListBal(1 To GYrCnt) As Double
              GLateListBal(GYrCnt) = ThisLateListBal#
              GLateListTot = ThisLateListBal#
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPrincBal(y) = OldRound(GPrincBal(y) + ThisPrincBal#)
                  GPrincTot = OldRound(GPrincTot# + ThisPrincBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GAdvBal(y) = OldRound(GAdvBal(y) + ThisAdvBal#)
                  GAdvTot = OldRound(GAdvTot# + ThisAdvBal#)
                  GLateListBal(y) = OldRound(GLateListBal(y) + ThisLateListBal#)
                  GLateListTot = OldRound(GLateListTot# + ThisLateListBal#)
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
'                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#) 'corrected 3/7/08
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
'                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt1Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#) 'corrected 3/7/08
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPrincBal(1 To GYrCnt) As Double
                GPrincBal(GYrCnt) = ThisPrincBal#
                GPrincTot# = OldRound(GPrincTot# + ThisPrincBal#)
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                ReDim Preserve GAdvBal(1 To GYrCnt) As Double
                GAdvBal(GYrCnt) = ThisAdvBal#
                GAdvTot# = OldRound(GAdvTot# + ThisAdvBal#)
                ReDim Preserve GLateListBal(1 To GYrCnt) As Double
                GLateListBal(GYrCnt) = ThisLateListBal#
                GLateListTot# = OldRound(GLateListTot + ThisLateListBal#)
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
            TransCnt = TransCnt + 1
'            TotCustCount = TotCustCount + 1
            If LngLastCust <> CustRec Then
              LngLastCust = CustRec
              TotCustCount = TotCustCount + 1
            End If
            
            '                   0            1             2                 3                  4              5            6
            Print #RptHandle, Town; dlm; CustName; dlm; CustRec; dlm; TaxTrans.TaxYear; dlm; Balance#; dlm; TotCustCount; dlm; GBal; dlm;
            '                         7                8                 9                   10                  11                 12                 13
            Print #RptHandle, ThisPrincBal#; dlm; ThisIntBal#; dlm; ThisAdvBal#; dlm; ThisLateListBal#; dlm; ThisOpt1Bal#; dlm; ThisOpt2Bal#; dlm; ThisOpt3Bal#; dlm;
            '                     14             15            16              17                18            19             20
            Print #RptHandle, GPrincTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm; GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm;
            '                    21              22            23
            Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
            If UseOpt = "Y" Then
              '                    24                      25                        26
              Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm;
            Else
              '                 24       25           26
              Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm;
            End If
            '                    27            28              29              30               31
            Print #RptHandle, PropType; dlm; ThisPin; dlm; CustTotBal; dlm; OverPay; dlm; fpcmbTaxYear.Text
            
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  Close
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPrinc = GPrincBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldAdv = GAdvBal(Nextz)
    HoldLateList = GLateListBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPrincBal(Nextz) = GPrincBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GAdvBal(Nextz) = GAdvBal(Thisz)
    GLateListBal(Nextz) = GLateListBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPrincBal(Thisz) = HoldPrinc
    GIntBal(Thisz) = HoldInt
    GAdvBal(Thisz) = HoldAdv
    GLateListBal(Thisz) = HoldLateList
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  SubRptFile$ = "TAXRPTS\TXMSTBALSUBDET.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle

  If InStr(CStr(OverPay), "E") Then OverPay = 0
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      '                        0               1                2          3
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 1; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPrincBal(x); dlm; GIntBal(x); dlm; GAdvBal(x); dlm; GLateListBal(x); dlm;
      '                         8                 9                10               11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; Opt1Desc; dlm;
      '                        12            13
      Print #SubRptHandle, Opt2Desc; dlm; Opt3Desc
    Else
      '                        0               1                2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 0; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPrincBal(x); dlm; GIntBal(x); dlm; GAdvBal(x); dlm; GLateListBal(x); dlm;
      '                         8                 9                10                11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; Opt1Desc; dlm;
      '                        12            13
      Print #SubRptHandle, Opt2Desc; dlm; Opt3Desc
    End If
  Next x
  
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  
  arTaxMasterBalDet.Show
  
  Exit Sub
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
    
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    'If TaxTrans.TaxYear = 2006 Then Stop
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
    
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc: '1/29/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = TaxTrans.Revenue.Principle1Pd / SaveAmt
  Disc1 = Disc1 * TaxTrans.DiscAmt
  Disc2 = TaxTrans.Revenue.RevOpt1Pd / SaveAmt
  Disc2 = Disc2 * TaxTrans.DiscAmt
  Disc3 = TaxTrans.Revenue.RevOpt2Pd / SaveAmt
  Disc3 = Disc3 * TaxTrans.DiscAmt
  Disc4 = TaxTrans.Revenue.RevOpt3Pd / SaveAmt
  Disc4 = Disc4 * TaxTrans.DiscAmt
  Disc5 = TaxTrans.Revenue.LateListPd / SaveAmt
  Disc5 = Disc5 * TaxTrans.DiscAmt
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5) 'added 9/19/07
  If Dif <> TaxTrans.DiscAmt Then
   If Disc1 > 0 Or Disc5 > 0 Then
     TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Abs(Dif - TaxTrans.DiscAmt)) '+ Dif)
   ElseIf Disc2 > 0 Then
     TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
   ElseIf Disc3 > 0 Then
     TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
   ElseIf Disc4 > 0 Then
     TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
   End If
  End If
  DiscApplied = True
  
  Return


ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMasterBalList", "PrintGraphicsDet", Erl)
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

Private Sub PrintTextDet()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPrincBal As Double
  Dim ThisIntBal As Double
  Dim ThisAdvBal As Double
  Dim ThisLateListBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim FF$
  Dim Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ThisAcct$, ThatAcct$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim ThisOP$
  Dim TestBal#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '9/17/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  Dim Dif As Double '9/19/07
  TotCustCount = 1
  
'  'on error goto ERRORSTUFF
  
  MaxLines = 58
  FF$ = Chr(12)
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If

  RptFile$ = "TAXRPTS\TXMSTBALDET.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  GoSub PrintHeader
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPrincBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GAdvBal(1 To 1) As Double
  ReDim GLateListBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo Inactive 'SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
'    CustTotBal = GetCustBalance(CustRec, -1)

    ThatAcct = CStr(TaxCust.Acct)
    
'    If ThatAcct <> 1032 Then
'      GoTo SkipOtherAccts:
'    End If
    
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    
    CustTotBal = TestBal
    OP = TestBal
   
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      ReDim PrincBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim AdvBal(1 To 1) As Double
      ReDim LateListBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      ThisAcct = ""
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        Balance = 0 'added 7/19/06
        ThisPrincBal = 0 'added 1/17/07
        ThisIntBal = 0 'added 1/17/07
        ThisAdvBal = 0 'added 1/17/07
        ThisLateListBal = 0 'added 1/17/07
        ThisOpt1Bal = 0 'added 1/17/07
        ThisOpt2Bal = 0 'added 1/17/07
        ThisOpt3Bal = 0 'added 1/17/07
        If CustTotBal < 0 Then
          Balance# = CustTotBal
          ThisPrincBal = CustTotBal 'added 1/17/07
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          TaxTrans.LastTrans = 0
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc '1/16/07
          ThisPrincBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          ThisPrincBal = OldRound(ThisPrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisAdvBal = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          ThisLateListBal = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            If ThisAcct <> ThatAcct Then
              TotCustCount = TotCustCount + 1
              GoSub PrintCustHeader
              ThisAcct = ThatAcct
            End If
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.RealPin) = "-1" Then
              ThisPin = "-1"
              PropType = "MOCK"
            ElseIf QPTrim$(TaxTrans.RealPin) <> "0" And QPTrim$(TaxTrans.RealPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.RealPin)
              PropType = "REAL"
            ElseIf QPTrim$(TaxTrans.PersPin) <> "0" And QPTrim$(TaxTrans.PersPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.PersPin)
              PropType = "PERSONAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            TransCnt = TransCnt + 1
            If LineCnt >= MaxLines - 1 Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, "Prop Type: " + PropType; Tab(23); "Pin #: " + ThisPin; Tab(43); "Tax Year: " + Using$("###0", TaxTrans.TaxYear); Tab(59); "Balance: " + Using$("$###,##0.00", Balance#) '; dlm; TransCnt; dlm; GBal; dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); "Principal:    " + Using$("$##,##0.00", ThisPrincBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt1Desc); Tab(69); Using$("$##,##0.00", ThisOpt1Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Principal:    " + Using$("$##,##0.00", ThisPrincBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                 Print #RptHandle, FF$
                 GoSub PrintHeader
                 GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); "Interest:     " + Using$("$##,##0.00", ThisIntBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt2Desc); Tab(69); Using$("$##,##0.00", ThisOpt2Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Interest:     " + Using$("$##,##0.00", ThisIntBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); "Advertising:  " + Using$("$##,##0.00", ThisAdvBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt3Desc); Tab(69); Using$("$##.##0.00", ThisOpt3Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Advertising:  " + Using$("$##,##0.00", ThisAdvBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            Print #RptHandle, Tab(5); "Late Listing: " + Using$("$##,##0.00", ThisLateListBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPrincBal(1 To GYrCnt) As Double
              GPrincBal(GYrCnt) = ThisPrincBal#
              GPrincTot = ThisPrincBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              ReDim Preserve GAdvBal(1 To GYrCnt) As Double
              GAdvBal(GYrCnt) = ThisAdvBal#
              GAdvTot = ThisAdvBal#
              ReDim Preserve GLateListBal(1 To GYrCnt) As Double
              GLateListBal(GYrCnt) = ThisLateListBal#
              GLateListTot = ThisLateListBal#
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPrincBal(y) = OldRound(GPrincBal(y) + ThisPrincBal#)
                  GPrincTot = OldRound(GPrincTot# + ThisPrincBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GAdvBal(y) = OldRound(GAdvBal(y) + ThisAdvBal#)
                  GAdvTot = OldRound(GAdvTot# + ThisAdvBal#)
                  GLateListBal(y) = OldRound(GLateListBal(y) + ThisLateListBal#)
                  GLateListTot = OldRound(GLateListTot# + ThisLateListBal#)
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
'                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#) 'corrected 3/7/08
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
'                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt1Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#) 'corrected 3/7/08
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPrincBal(1 To GYrCnt) As Double
                GPrincBal(GYrCnt) = ThisPrincBal#
                GPrincTot# = OldRound(GPrincTot# + ThisPrincBal#)
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                ReDim Preserve GAdvBal(1 To GYrCnt) As Double
                GAdvBal(GYrCnt) = ThisAdvBal#
                GAdvTot# = OldRound(GAdvTot# + ThisAdvBal#)
                ReDim Preserve GLateListBal(1 To GYrCnt) As Double
                GLateListBal(GYrCnt) = ThisLateListBal#
                GLateListTot# = OldRound(GLateListTot + ThisLateListBal#)
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
SkipOtherAccts:
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPrinc = GPrincBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldAdv = GAdvBal(Nextz)
    HoldLateList = GLateListBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPrincBal(Nextz) = GPrincBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GAdvBal(Nextz) = GAdvBal(Thisz)
    GLateListBal(Nextz) = GLateListBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPrincBal(Thisz) = HoldPrinc
    GIntBal(Thisz) = HoldInt
    GAdvBal(Thisz) = HoldAdv
    GLateListBal(Thisz) = HoldLateList
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  If InStr(CStr(OverPay), "E") Then OverPay = 0
  Print #RptHandle, FF$
  GoSub PrintEndHeader
  Print #RptHandle, "Total Entries: " + Using$("####0", TotCustCount)
  Print #RptHandle, "Total Tax Balance: "; Tab(30); Using$("$###,###,##0.00", GBal)
  Print #RptHandle,
  Print #RptHandle, "Principal Total:          "; Tab(30); Using$("$###,###,##0.00", GPrincTot)
  Print #RptHandle, "Interest Total:           "; Tab(30); Using$("$###,###,##0.00", GIntTot)
  Print #RptHandle, "Advertising Total:        "; Tab(30); Using$("$###,###,##0.00", GAdvTot)
  Print #RptHandle, "Late Listing Total:       "; Tab(30); Using$("$###,###,##0.00", GLateListTot)
  If fpcmbTaxYear.Text = "All" Then
    Print #RptHandle, "OverPayments:             "; Tab(30); Using$("$###,###,##0.00", OverPay)
  Else
    Print #RptHandle,
  End If
  LineCnt = LineCnt + 8
  If QPTrim$(Opt1Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt1Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt1Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(Opt2Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt2Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt2Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(Opt3Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt3Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt3Tot)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 2
  
  For x = 1 To GYrCnt
    Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed"
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    If x = GYrCnt Then
      Print #RptHandle, Tab(2); Using$("###0", GYears(x)); Tab(30); Using$("$###,###,##0.00", GYearBal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, Tab(2); "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    Else
      Print #RptHandle, Using$("###0", GYears(x)); Tab(30); Using$("$###,###,##0.00", GYearBal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    Print #RptHandle, Tab(5); "Principal:"; Tab(30); Using$("$###,###,##0.00", GPrincBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Interest: "; Tab(30); Using$("$###,###,##0.00", GIntBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Advertising:"; Tab(30); Using$("$###,###,##0.00", GAdvBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Late Listing:"; Tab(30); Using$("$###,###,##0.00", GLateListBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    If QPTrim$(Opt1Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt1Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt1Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(Opt2Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt2Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt2Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(Opt3Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt3Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt3Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    Print #RptHandle,
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed"; Tab(60); "Over Payments"
    End If
  Next x
  
  Print #RptHandle, FF$
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  
  'MsgBox (CStr(TotCustCount))
  
  ViewPrint RptFile, "Tax Master Balance Listing", True
  
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Detail"
  Print #RptHandle, Town; Tab(36); "Year: " + fpcmbTaxYear.Text
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If
  If ActiveFlag = "B" Then
    Print #RptHandle, "Customer Status: " + "Active And Inactive"; Tab(65); "(I) = Inactive"
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  End If
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(59); "Tax Year"; Tab(73); "Balance"
  Print #RptHandle, String(79, "-")
  LineCnt = 6
  
  Return
  
PrintCustHeader:
  If LineCnt <> 6 Then
    Print #RptHandle, String(79, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("####0", CustRec); Tab(8); CustName; Tab(55); "Cust Balance: " + QPTrim$(Using$("$###,###,##0.00", CustTotBal))
  Print #RptHandle, String(79, "-")
  LineCnt = LineCnt + 2
  Return

PrintEndHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Detail"
  Print #RptHandle, Town
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If

  Print #RptHandle, String(79, "-")
  LineCnt = 4
  
  Return
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
    
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc: '1/29/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = TaxTrans.Revenue.Principle1Pd / SaveAmt
  Disc1 = Disc1 * TaxTrans.DiscAmt
  Disc2 = TaxTrans.Revenue.RevOpt1Pd / SaveAmt
  Disc2 = Disc2 * TaxTrans.DiscAmt
  Disc3 = TaxTrans.Revenue.RevOpt2Pd / SaveAmt
  Disc3 = Disc3 * TaxTrans.DiscAmt
  Disc4 = TaxTrans.Revenue.RevOpt3Pd / SaveAmt
  Disc4 = Disc4 * TaxTrans.DiscAmt
  Disc5 = TaxTrans.Revenue.LateListPd / SaveAmt
  Disc5 = Disc5 * TaxTrans.DiscAmt
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5) 'added 9/19/07
  If Dif <> TaxTrans.DiscAmt Then
   If Disc1 > 0 Or Disc5 > 0 Then
     TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Abs(Dif - TaxTrans.DiscAmt)) '+ Dif)
   ElseIf Disc2 > 0 Then
     TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
   ElseIf Disc3 > 0 Then
     TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
   ElseIf Disc4 > 0 Then
     TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
   End If
  End If
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMasterBalList", "PrintTextDet", Erl)
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

Private Function GetOverPayAmount(TheYear As Integer) As Double
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim OP As Double
  Dim TotOP
  Dim ActiveFlag$
  Dim TestBal#
  Dim ThisRec As Long
  Dim TTHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTTRecs As Long
  
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  OpenTaxCustFile CHandle, NumOfCRecs
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show , Me
  frmTaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    If TheYear < 0 Then
'      OP = OldRound(GetCustBalance(x, -1))
      GoSub GetAllBalance
      OP = TestBal
      OP = OldRound(GetCustBalance(x, -1))
      If OP < 0 Then
        TotOP = OldRound(TotOP + Abs(OP))
      End If
    Else
'      OP = OldRound(GetCustBalanceForYear(x, TheYear))
      GoSub GetYearBalance
      OP = TestBal#
      If OP < 0 Then
        TotOP = OldRound(TotOP + Abs(OP))
      End If
    End If
Inactive:
    frmTaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Function
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  Close CHandle
  Close TTHandle
  GetOverPayAmount = TotOP
  
  Exit Function
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TheYear <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo SkipItYear
    End If
   If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
End Function
