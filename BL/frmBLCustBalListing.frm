VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustBalListing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Balance Listing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCustBalListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6516
      Left            =   1920
      TabIndex        =   8
      Top             =   1020
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11493
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
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLCustBalListing.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2970
         TabIndex        =   0
         Tag             =   $"frmBLCustBalListing.frx":08E6
         Top             =   1680
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
         ColDesigner     =   "frmBLCustBalListing.frx":0992
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2970
         TabIndex        =   4
         Tag             =   $"frmBLCustBalListing.frx":0C89
         Top             =   4410
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
         ColDesigner     =   "frmBLCustBalListing.frx":0D42
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   5085
         TabIndex        =   3
         Tag             =   $"frmBLCustBalListing.frx":1039
         Top             =   3750
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
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
         ColDesigner     =   "frmBLCustBalListing.frx":10D2
      End
      Begin LpLib.fpCombo fpcmbBalances 
         Height          =   405
         Left            =   5235
         TabIndex        =   2
         Tag             =   $"frmBLCustBalListing.frx":13C9
         Top             =   3075
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
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
         ColDesigner     =   "frmBLCustBalListing.frx":1482
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3210
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5325
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmBLCustBalListing.frx":1779
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5235
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLCustBalListing.frx":1957
         Top             =   5325
         Width           =   1830
         _Version        =   131072
         _ExtentX        =   3228
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
         ButtonDesigner  =   "frmBLCustBalListing.frx":1A02
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   390
         Left            =   4845
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmBLCustBalListing.frx":1BE1
         Top             =   2355
         Width           =   1815
         _Version        =   131072
         _ExtentX        =   3201
         _ExtentY        =   688
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
         ButtonDesigner  =   "frmBLCustBalListing.frx":1D28
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   396
         Left            =   2976
         TabIndex        =   1
         Tag             =   $"frmBLCustBalListing.frx":1F0C
         Top             =   2352
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
      Begin fpBtnAtlLibCtl.fpBtn fpcmdHelp 
         Height          =   645
         Left            =   915
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   $"frmBLCustBalListing.frx":20DF
         Top             =   5325
         Width           =   2175
         _Version        =   131072
         _ExtentX        =   3836
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
         ButtonDesigner  =   "frmBLCustBalListing.frx":21AF
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Zero Balances (Y/N)?:"
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
         Left            =   1632
         TabIndex        =   17
         Top             =   3168
         Width           =   3372
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
         Left            =   960
         TabIndex        =   15
         Top             =   6000
         Width           =   2100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Inactive Accounts?:"
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
         Left            =   1824
         TabIndex        =   13
         Top             =   3840
         Width           =   3036
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Balance Listing"
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
         Height          =   396
         Left            =   2016
         TabIndex        =   12
         Top             =   576
         Width           =   3948
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
         Left            =   1392
         TabIndex        =   11
         Top             =   1776
         Width           =   1308
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
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
         Left            =   1344
         TabIndex        =   10
         Top             =   2448
         Width           =   1356
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3708
         Left            =   1008
         Top             =   1392
         Width           =   5964
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
         Height          =   348
         Left            =   1200
         TabIndex        =   9
         Top             =   4512
         Width           =   1500
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   1872
      TabIndex        =   16
      Top             =   7872
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
      MaxWidth        =   3000
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
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6780
      Left            =   1824
      Top             =   888
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLCustBalListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
  DoEvents
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
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
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdCodeList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call fpcmdHelp_Click
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
      KillFile "custbalList.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustBalListing.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  
  lblBalloon.Visible = False
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  fptxtCatCode.ToolTipText = "You can select ALL or you can select a specific category for which to print this report."
'  cmdCodeList.ToolTipText = "Press to bring up a complete category list."
'  fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include active and inactive accounts."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  fpcmdHelp.ToolTipText = "Press this button to activate informational balloons. Then press it again to deactivate balloons."
'  cmdExit.ToolTipText = "Press to exit this screen and return to the main Customer Reports menu."
'  cmdProcess.ToolTipText = "Press to generate this report for printing."
  One = 1
  DHandle = FreeFile
  Open "custbalList.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fptxtCatCode.Text = "ALL"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbIncInactive.Text = "No"
  fpcmbIncInactive.AddItem "Yes"
  fpcmbIncInactive.AddItem "No"
  fpcmbBalances.Text = "No"
  fpcmbBalances.AddItem "Yes"
  fpcmbBalances.AddItem "No"
End Sub

Private Sub fpcmbBalances_Change()
  If QPTrim$(fpcmbBalances.Text) = "" Then
    fpcmbBalances.Text = "No"
  End If
End Sub

Private Sub fpcmbBalances_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBalances.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBalances.ListIndex = -1
  End If
  If fpcmbBalances.ListDown <> True Then
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

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcmbPrintOrder_Change()
  If QPTrim$(fpcmbPrintOrder.Text) = "" Then
    fpcmbPrintOrder.Text = "Billing Name Order"
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtCatCode.SetFocus
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
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
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

Private Sub cmdExit_Click()
  frmBLCustReportsMenu.Show
  KillFile "custbalList.dat"
  DoEvents
  Unload frmBLCustBalListing
End Sub

Private Sub cmdProcess_Click()

  If Check4ValidCatNum(QPTrim$(fptxtCatCode.Text)) = False Then
    frmBLMessageBoxJr.Label1.Caption = "The category code number entered is not valid. Please enter a valid category code number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    If fptxtCatCode.Enabled = True Then
      fptxtCatCode.SetFocus
    End If
    Exit Sub
  End If
  
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcmbPrintOpt.Text = "Text" Then
    frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim IdxCnt As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim CustCnt As Integer
  Dim NumOfCustRecs As Integer
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim cnt As Integer
  Dim CustomerNumber As Integer
  Dim TCat$, TotalBal#
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim InActiveFlag As Boolean
  Dim ShowZBalFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  ShowZBalFlag = True
  
  If fpcmbBalances.Text = "No" Then
    ShowZBalFlag = False
  End If
  
  InActiveFlag = False
  
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  Else
    NumFlag = True
    NameFlag = False
  End If
  TCat = QPTrim$(fptxtCatCode.Text)
  ReportFile$ = "ARCusBal.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintCustBalRptHeader
  
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)

  If NameFlag = True Then
    OpenCustNameIdxFile IdxHandle
    IdxCnt = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    IdxCnt = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If IdxCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim ThisIdx(1 To IdxCnt) As Integer
  
  If NameFlag = True Then
    For x = 1 To IdxCnt
      Get IdxHandle, x, CustNameIdxRec
      ThisIdx(x) = CustNameIdxRec.CustRec
    Next x
  Else
    For x = 1 To IdxCnt
      Get IdxHandle, x, CustNumIdxRec
      ThisIdx(x) = CustNumIdxRec.CustRec
    Next x
  End If
  
  Close IdxHandle
  
  frmBLShowPctComp.Label1 = "Loading Customer Balance Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For cnt = 1 To IdxCnt
    Get CHandle, ThisIdx(cnt), CustRec
    If ShowZBalFlag = False And CustRec.AcctBal = 0 Then GoTo Inactive
    If InActiveFlag = False Then
      If QPTrim$(CustRec.Inactive) = "Y" Then
        GoTo Inactive
      End If
    End If
    CustomerNumber = ThisIdx(cnt)
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      If TCat$ = "ALL" Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.AcctBal)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.AcctBal)
        LineCnt = LineCnt + 1
      ElseIf TCat$ = "PENALTIES" Then
        If CustRec.PenBal > 0 Then
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintCustBalRptHeader
          End If
          Print #RptHandle, Using("####0", CustomerNumber);
          Print #RptHandle, Tab(10); CustRec.BillName;
          Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.PenBal)
          CustCnt = CustCnt + 1
          TotalBal# = OldRound#(TotalBal# + CustRec.PenBal)
          LineCnt = LineCnt + 1
        End If
      ElseIf TCat$ = "ISSUANCE" Then
        If CustRec.IssuanceBal > 0 Then
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintCustBalRptHeader
          End If
          Print #RptHandle, Using("####0", CustomerNumber);
          Print #RptHandle, Tab(10); CustRec.BillName;
          Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.IssuanceBal)
          CustCnt = CustCnt + 1
          TotalBal# = OldRound#(TotalBal# + CustRec.IssuanceBal)
          LineCnt = LineCnt + 1
        End If
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT1) Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.FeeLicBal1)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal1)
        LineCnt = LineCnt + 1
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT2) Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.FeeLicBal2)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal2)
        LineCnt = LineCnt + 1
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT3) Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.FeeLicBal3)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal3)
        LineCnt = LineCnt + 1
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT4) Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.FeeLicBal4)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal4)
        LineCnt = LineCnt + 1
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT5) Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintCustBalRptHeader
        End If
        Print #RptHandle, Using("####0", CustomerNumber);
        Print #RptHandle, Tab(10); CustRec.BillName;
        Print #RptHandle, Tab(68); Using("$###,###,0.00", CustRec.FeeLicBal5)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal5)
        LineCnt = LineCnt + 1
      End If
    End If
Inactive:
    frmBLShowPctComp.ShowPctComp cnt, IdxCnt
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      fpcmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  GoSub PrintCustBalRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  If CustCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers with balances for the parameters entered."
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Customer Balance Listing", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Customer Balance report processed for category " + QPTrim$(fptxtCatCode.Text) + " in text format.")
  Exit Sub
  
PrintCustBalRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License : Customer Balance Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "Category: "; TCat$ + "/" + GetCatDesc(TCat$)
  Print #RptHandle, "Cust #"; Tab(10); "Billing Name"; Tab(66); "Account Balance"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return
  
PrintCustBalRptEnding:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, "Total Customers Printed: "; Using("####0", CustCnt);
  Print #RptHandle, Tab(68); Using("$#,###,##0.00", TotalBal#)
  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustBalListing", "PrintText", Erl)
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

Private Sub fpcmdHelp_Click()
  If InStr(fpcmdHelp.Text, "On") Then
    fpcmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpcmbPrintOrder.ToolTipText = ""
    fptxtCatCode.ToolTipText = ""
    cmdCodeList.ToolTipText = ""
    fpcmbIncInactive.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    fpcmdHelp.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    fptxtCatCode.ToolTipText = "You can select ALL or you can select a specific category for which to print this report."
'    cmdCodeList.ToolTipText = "Press to bring up a complete category list."
'    fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include active and inactive accounts."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    fpcmdHelp.ToolTipText = "Press this button to activate informational balloons. Then press it again to deactivate balloons."
'    cmdExit.ToolTipText = "Press to exit this screen and return to the main Customer Reports menu."
'    cmdProcess.ToolTipText = "Press to generate this report for printing."
  End If
End Sub

Private Sub fptxtCatCode_Change()
  If QPTrim$(fptxtCatCode.Text) = "" Then
    fptxtCatCode.Text = "ALL"
  End If
  If QPTrim$(fptxtCatCode.Text) = "P" Or QPTrim$(fptxtCatCode.Text) = "p" Then
    fptxtCatCode.Text = "PENALTIES"
  End If
  If QPTrim$(fptxtCatCode.Text) = "I" Or QPTrim$(fptxtCatCode.Text) = "i" Then
    fptxtCatCode.Text = "ISSUANCE"
  End If
  
End Sub

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim x As Integer
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim IdxCnt As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim CustCnt As Integer
  Dim NumOfCustRecs As Integer
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim cnt As Integer
  Dim CustomerNumber As Integer
  Dim TCat$, TotalBal#
  Dim RptHandle As Integer
  Dim dlm$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim TownName$
  Dim InActiveFlag As Boolean
  Dim ShowZBalFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  ShowZBalFlag = True
  
  If fpcmbBalances.Text = "No" Then
    ShowZBalFlag = False
  End If
  
  InActiveFlag = False
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName$ = QPTrim$(TownRec.TownName)
  
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  Else
    NumFlag = True
    NameFlag = False
  End If
  
  TCat = QPTrim$(fptxtCatCode.Text)
  ReportFile$ = "BLRPTS\ARCusBal.Rpt"
  CustCnt = 0
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)

  If NameFlag = True Then
    OpenCustNameIdxFile IdxHandle
    IdxCnt = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    IdxCnt = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If IdxCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim ThisIdx(1 To IdxCnt) As Integer
  
  If NameFlag = True Then
    For x = 1 To IdxCnt
      Get IdxHandle, x, CustNameIdxRec
      ThisIdx(x) = CustNameIdxRec.CustRec
    Next x
  Else
    For x = 1 To IdxCnt
      Get IdxHandle, x, CustNumIdxRec
      ThisIdx(x) = CustNumIdxRec.CustRec
    Next x
  End If
  
  Close IdxHandle
  
  frmBLShowPctComp.Label1 = "Loading Customer Balance Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For cnt = 1 To IdxCnt
    Get CHandle, ThisIdx(cnt), CustRec
    If ShowZBalFlag = False And CustRec.AcctBal = 0 Then GoTo Inactive
    If InActiveFlag = False Then
      If QPTrim$(CustRec.Inactive) = "Y" Then
        GoTo Inactive
      End If
    End If
    CustomerNumber = ThisIdx(cnt)
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      If TCat$ = "ALL" Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.AcctBal; dlm; TCat$
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.AcctBal)
      ElseIf TCat$ = "PENALTIES" Then
        If CustRec.PenBal > 0 Then
          Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.PenBal; dlm; TCat$ + "/" + GetCatDesc(TCat$)
          CustCnt = CustCnt + 1
          TotalBal# = OldRound#(TotalBal# + CustRec.PenBal)
        End If
      ElseIf TCat$ = "ISSUANCE" Then
        If CustRec.IssuanceBal > 0 Then
          Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.IssuanceBal; dlm; TCat$ + "/" + GetCatDesc(TCat$)
          CustCnt = CustCnt + 1
          TotalBal# = OldRound#(TotalBal# + CustRec.IssuanceBal)
        End If
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT1) Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.FeeLicBal1; dlm; TCat$ + "/" + GetCatDesc(TCat$)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal1)
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT2) Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.FeeLicBal2; dlm; TCat$ + "/" + GetCatDesc(TCat$)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal2)
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT3) Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.FeeLicBal3; dlm; TCat$ + "/" + GetCatDesc(TCat$)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal3)
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT4) Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.FeeLicBal4; dlm; TCat$ + "/" + GetCatDesc(TCat$)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal4)
      ElseIf TCat$ = QPTrim$(CustRec.BILLCAT5) Then
        Print #RptHandle, TownName$; dlm; CustomerNumber; dlm; CustRec.BillName; dlm; CustRec.FeeLicBal5; dlm; TCat$ + "/" + GetCatDesc(TCat$)
        CustCnt = CustCnt + 1
        TotalBal# = OldRound#(TotalBal# + CustRec.FeeLicBal5)
      End If
    End If
Inactive:
    frmBLShowPctComp.ShowPctComp cnt, IdxCnt
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      fpcmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt
  Close         'Close all open files now
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If CustCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers with balances for the parameters entered."
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLCustBalRpt.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Customer Balance report processed for category " + QPTrim$(fptxtCatCode.Text) + " in graphics format.")
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustBalListing", "PrintGraphics", Erl)
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

Private Sub fptxtCatCode_LostFocus()
  If Not IsNumeric(fptxtCatCode.Text) And fptxtCatCode.Text <> "PENALTIES" And fptxtCatCode.Text <> "ISSUANCE" Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub
