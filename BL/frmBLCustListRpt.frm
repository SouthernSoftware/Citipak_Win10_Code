VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustListRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Detailed Customer Listing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCustListRpt.frx":0000
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
      TabIndex        =   6
      Top             =   1176
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
      Picture         =   "frmBLCustListRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   0
         Tag             =   $"frmBLCustListRpt.frx":08E6
         Top             =   1836
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLCustListRpt.frx":0992
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   4
         Tag             =   $"frmBLCustListRpt.frx":0C8D
         Top             =   4368
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLCustListRpt.frx":0D46
      End
      Begin LpLib.fpCombo fpcmbFees 
         Height          =   384
         Left            =   4656
         TabIndex        =   2
         Tag             =   $"frmBLCustListRpt.frx":1041
         Top             =   3168
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLCustListRpt.frx":1152
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   384
         Left            =   5040
         TabIndex        =   3
         Tag             =   "You can elect to include all inactive accounts on this report. Select 'Yes' in the drop down list to include inactive accounts."
         Top             =   3756
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLCustListRpt.frx":144D
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3315
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5370
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
         ButtonDesigner  =   "frmBLCustListRpt.frx":1748
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5190
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLCustListRpt.frx":1926
         Top             =   5370
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
         ButtonDesigner  =   "frmBLCustListRpt.frx":19D1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   405
         Left            =   4605
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmBLCustListRpt.frx":1BB0
         Top             =   2490
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   714
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
         ButtonDesigner  =   "frmBLCustListRpt.frx":1C69
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   390
         Left            =   2730
         TabIndex        =   1
         Tag             =   $"frmBLCustListRpt.frx":1E4D
         Top             =   2490
         Width           =   1830
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
         Left            =   1005
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   $"frmBLCustListRpt.frx":1F64
         Top             =   5370
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
         ButtonDesigner  =   "frmBLCustListRpt.frx":2034
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
         Left            =   1056
         TabIndex        =   16
         Top             =   6048
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
         Height          =   348
         Left            =   1776
         TabIndex        =   14
         Top             =   3840
         Width           =   3036
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show Current Fees?:"
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
         Left            =   2112
         TabIndex        =   13
         Top             =   3264
         Width           =   2316
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
         Left            =   1152
         TabIndex        =   12
         Top             =   4464
         Width           =   1500
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
         Left            =   1155
         TabIndex        =   11
         Top             =   2595
         Width           =   1350
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
         Caption         =   "Customer Detail Listing"
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
         Left            =   1776
         TabIndex        =   10
         Top             =   576
         Width           =   4332
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3612
         Left            =   1008
         Top             =   1488
         Width           =   5964
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   2040
      TabIndex        =   17
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
      Left            =   1800
      Top             =   1044
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLCustListRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim TextPrintOpt As Integer

Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  TextPrintOpt = 0
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
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustListRpt.")
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
'  fpcmbFees.ToolTipText = "This report can be printed with each customer's current business license fee (not outstanding balance)."
'  fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include all."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  cmdExit.ToolTipText = "Press this button to exit to 'Business License Reports' menu."
'  cmdProcess.ToolTipText = "Press to activate this report."
  One = 1
  DHandle = FreeFile
  Open "custlistRpt.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fptxtCatCode.Text = "ALL"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbFees.Text = "No"
  fpcmbFees.AddItem "Yes"
  fpcmbFees.AddItem "No"
  fpcmbIncInactive.Text = "No"
  fpcmbIncInactive.AddItem "Yes"
  fpcmbIncInactive.AddItem "No"

End Sub

Private Sub fpcmbFees_Change()
  If QPTrim$(fpcmbFees.Text) = "" Then
    fpcmbFees.Text = "No"
  End If
End Sub

Private Sub fpcmbFees_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbFees.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFees.ListIndex = -1
  End If
  If fpcmbFees.ListDown <> True Then
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
  KillFile "custlistRpt.dat"
  DoEvents
  Unload frmBLCustListRpt
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

Private Sub fpcmdHelp_Click()
  If InStr(fpcmdHelp.Text, "On") Then
    fpcmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpcmbPrintOrder.ToolTipText = ""
    fptxtCatCode.ToolTipText = ""
    cmdCodeList.ToolTipText = ""
    fpcmbFees.ToolTipText = ""
    fpcmbIncInactive.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    fptxtCatCode.ToolTipText = "You can select ALL or you can select a specific category for which to print this report."
'    cmdCodeList.ToolTipText = "Press to bring up a complete category list."
'    fpcmbFees.ToolTipText = "This report can be printed with each customer's current business license fee (not outstanding balance)."
'    fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include all."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    cmdExit.ToolTipText = "Press this button to exit to 'Business License Reports' menu."
'    cmdProcess.ToolTipText = "Press to activate this report."
  End If
End Sub

Private Sub fptxtCatCode_Change()
  If QPTrim$(fptxtCatCode.Text) = "" Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim TCat$, CustNum$
  Dim ShowFee As Boolean
  Dim CustFee#, FeeAmt#, cnt As Double
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim Prorate#
  Dim CatCode$, Snt&, Mult#
  Dim Revenue#
  Dim InActiveFlag As Boolean
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisCustCnt As Double
  Dim TownName$
  Dim ThisCat$
  Dim Nextx As Double
  Dim NextCust As Double
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName)
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  TCat$ = QPTrim$(fptxtCatCode.Text)
  InActiveFlag = False
  
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  
  If fpcmbFees.Text = "No" Then
    ShowFee = False
  Else
    ShowFee = True
  End If
  
  ReportFile$ = "ARDetCus.PRN"
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  OpenCatCodeFile CHandle
  NumOfARCatRecs = LOF(CHandle) \ Len(CodeRec)
  
  If NumOfARCatRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "There are no category codes on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
'  If TextPrintOpt = 2 Then GoTo Print2
'  GoSub PrintDetailCustomerRptHeader
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
    NameFlag = False
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection for Print Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  If NameFlag = True Then
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If NumOfCustIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCustFile CustHandle
  
  ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
  
  DoEvents
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
      For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
'  If TextPrintOpt = 2 Then GoTo Print2
  
  GoSub PrintDetailCustomerRptHeader
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For cnt = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(cnt), CustRec
    If InActiveFlag = False Then
      If QPTrim$(CustRec.Inactive) = "Y" Then
        GoTo Inactive
      End If
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      If TCat$ = QPTrim$(CustRec.BILLCAT1) Or TCat$ = QPTrim$(CustRec.BILLCAT2) Or TCat$ = QPTrim$(CustRec.BILLCAT3) Or TCat$ = QPTrim$(CustRec.BILLCAT4) Or TCat$ = QPTrim$(CustRec.BILLCAT5) Or TCat$ = "ALL" Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintDetailCustomerRptHeader
        End If
        If ShowFee Then
          GoSub GetCustFee
        End If
        CustNum$ = Space$(6)
        RSet CustNum$ = QPTrim$(CustRec.CustNumb)
        Print #RptHandle, QPTrim$(CustNum$);
        If QPTrim$(CustRec.Inactive) = "Y" Then
          Print #RptHandle, Tab(13); "Inactive";
        Else
          Print #RptHandle, Tab(13); "Active";
        End If
        Print #RptHandle, Tab(25); Using("##0.00", CustRec.Prorate);
        Print #RptHandle, Tab(39); "Lic # "; QPTrim$(CustRec.LICENSE); Tab(60); "Valid to: "; MakeRegDate(CustRec.VALID) 'end line 1
        Print #RptHandle, QPTrim$(CustRec.BillName); Tab(40); "Category: "; QPTrim$(CustRec.BILLCAT1); " / "; QPTrim$(CustRec.BILLCAT2); " / "; QPTrim$(CustRec.BILLCAT3); " / "; QPTrim$(CustRec.BILLCAT4); " / "; QPTrim$(CustRec.BILLCAT5) 'end line 2
        If QPTrim$(CustRec.WPHONE) = "(" Then CustRec.WPHONE = ""
        Print #RptHandle, RTrim$(CustRec.ADDRESS1); Tab(47); " Work Phone: "; QPTrim$(CustRec.WPHONE) 'end line 3
        Print #RptHandle, RTrim$(CustRec.ADDRESS2) 'end line 4
        Print #RptHandle, RTrim$(CustRec.City); ", "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode);
        If ShowFee Then
          Print #RptHandle, Tab(50); "Lic Fees: $"; Using("###,##0.00", CustFee#) 'end line 5
        Else
          Print #RptHandle, 'end line 5
        End If
        Print #RptHandle, String$(80, "-") 'end line 6
        CustCnt = CustCnt + 1
        LineCnt = LineCnt + 6
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustIdxRecs
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
Inactive:
  Next cnt
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  GoSub PrintDetailCustomerRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  If CustCnt = 0 Then
    Close         'Close all open files now
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers on file that fit the criteria entered."
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Detailed Customer Listing", True
  
  KillFile ReportFile$
  
  MainLog ("The 'Customer Detail Report' was processed in text format.")
  
  Exit Sub
  
  
PrintDetailCustomerRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License: Detailed Customer Listing"
  Print #RptHandle, TownName
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Str(Page)
  If ShowFee = True Then
    If TownRec.IssFee > 0 Then
      If TCat$ = "ALL" Then
        Print #RptHandle, "Category: " + TCat$; Tab(46); "Fees include a " + QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) + " issuance fee."
      Else
        Print #RptHandle, "Category: " + TCat$ + "/" + GetCatDesc(TCat$); Tab(46); "Fees include a " + QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) + " issuance fee."
      End If
    Else
      If TCat$ = "ALL" Then
        Print #RptHandle, "Category: " + TCat$
      Else
        Print #RptHandle, "Category: " + TCat$ + "/" + GetCatDesc(TCat$)
      End If
    End If
  Else
    If TCat$ = "ALL" Then
      Print #RptHandle, "Category: " + TCat$
    Else
      Print #RptHandle, "Category: " + TCat$ + "/" + GetCatDesc(TCat$)
    End If
  End If
  
  Print #RptHandle, "Acct No   Inactive     Prorate         Lic No"
  Print #RptHandle, String$(80, "=")
  LineCnt = 5
  Return
  
PrintDetailCustomerRptEnding:
  Print #RptHandle, "Total Customers Printed: "; Using("####0", CustCnt)
  Print #RptHandle,
  If TCat$ <> "ALL" Then
    Print #RptHandle, FF$
  End If
  Return
  
GetCustFee:
  CustFee# = 0
  FeeAmt# = 0
  Prorate# = CustRec.Prorate
  If Prorate# >= 100 Or Prorate# < 0 Then
    Prorate# = 1
  Else
    Prorate# = OldRound#(Prorate# * 0.01)
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        CustRec.DESC1 = CodeRec.CODEDESC           'Reset Code Descriptions
        If CodeRec.CodeType = "F" Then
          FeeAmt# = OldRound(CodeRec.Fee * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV1
          FeeAmt# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV1
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C2
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C2:             'Catagory #2
  
  CustFee# = OldRound#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt# = CodeRec.Fee
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV2
          FeeAmt# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV2
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo C3
          End If
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C3:
  CustFee# = OldRound#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt# = CodeRec.Fee
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV3
          FeeAmt# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV3
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c4
          End If
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 3
  
c4:
  CustFee# = OldRound#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt# = CodeRec.Fee
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV4
          FeeAmt# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV4
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo c5
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
c5:
  CustFee# = OldRound#(CustFee# + FeeAmt#)
  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt# = CodeRec.Fee
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV5
          FeeAmt# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt# = OldRound(FeeAmt# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV5
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
            FeeAmt# = OldRound(FeeAmt# * Prorate#)
            GoTo SkipEm
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
SkipEm:
  CustFee# = OldRound#(CustFee# + FeeAmt# + TownRec.IssFee)
  FeeAmt# = 0
  
  Return
  

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustListRpt", "PrintText", Erl)
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

Private Sub PrintGraphics()
  Dim ReportFile$, ReportFileSub$
  Dim x As Double
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim TCat$, CustNum$
  Dim ShowFee As Boolean
  Dim CustFee#, FeeAmt#, cnt As Double
  Dim RptHandle As Integer, RptHandleSub As Integer
  Dim Prorate#, ThisCustCnt As Double
  Dim CatCode$, Snt&, Mult#
  Dim Revenue#, ThisCat$
  Dim dlm$, Nextx As Double
  Dim TownName$, NextCust As Double
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim InActiveFlag As Boolean
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  InActiveFlag = False
  
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName)
  
  If fpcmbFees.Text = "No" Then
    ShowFee = False
  Else
    ShowFee = True
  End If
  
  ReportFile$ = "BLRPTS\ARDetCus.RPT"
  CustCnt = 0
  
  OpenCatCodeFile CHandle
  NumOfARCatRecs = LOF(CHandle) \ Len(CodeRec)
  
  If NumOfARCatRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "There are no category codes on file."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
    NameFlag = False
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection for Print Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  If NameFlag = True Then
'    OpenSrchNameIdxFile IdxHandle
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If NumOfCustIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCustFile CustHandle
  
  ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  DoEvents
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
      For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
  For cnt = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(cnt), CustRec
    If InActiveFlag = False Then
      If QPTrim$(CustRec.Inactive) = "Y" Then
        GoTo Inactive
      End If
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
      TCat$ = QPTrim$(fptxtCatCode.Text)
      'check to see if this customer has a category code
      'that fits the user's criteria
      If TCat$ = QPTrim$(CustRec.BILLCAT1) Or TCat$ = QPTrim$(CustRec.BILLCAT2) Or TCat$ = QPTrim$(CustRec.BILLCAT3) Or TCat$ = QPTrim$(CustRec.BILLCAT4) Or TCat$ = QPTrim$(CustRec.BILLCAT5) Or TCat$ = "ALL" Then
        'if the user elects to print fees then the program
        'activates GetCustFees which calculates license fees
        'but does not consider outstanding balances or issuance
        'fees
        If ShowFee Then
          GoSub GetCustFee
        End If
        CustNum$ = Space$(6)
        RSet CustNum$ = QPTrim$(CustRec.CustNumb)
        Print #RptHandle, TownName; dlm; QPTrim$(CustNum$); dlm;
        If QPTrim$(CustRec.Inactive) = "Y" Then
          Print #RptHandle, "Inactive"; dlm; CustRec.Prorate / 100; dlm;
        Else
          Print #RptHandle, "Active"; dlm; CustRec.Prorate / 100; dlm;
        End If
        Print #RptHandle, QPTrim$(CustRec.LICENSE); dlm; MakeRegDate(CustRec.VALID); dlm;
        Print #RptHandle, QPTrim$(CustRec.BillName); dlm; QPTrim$(CustRec.BILLCAT1) + "/" + QPTrim$(CustRec.BILLCAT2) + "/" + QPTrim$(CustRec.BILLCAT3) + "/" + QPTrim$(CustRec.BILLCAT4) + "/" + QPTrim$(CustRec.BILLCAT5); dlm;
        Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm;
        If Mid(CustRec.WPHONE, 2, 1) <> " " Then
          Print #RptHandle, QPTrim$(CustRec.WPHONE); dlm;
        Else
          Print #RptHandle, ""; dlm;
        End If
        Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm; QPTrim$(CustRec.City); dlm; QPTrim$(CustRec.State); dlm; QPTrim$(CustRec.ZipCode); dlm;
        If ShowFee Then
          Print #RptHandle, CustFee#; dlm;
        Else
          Print #RptHandle, "    "; dlm;
        End If
        If ShowFee = False Then TownRec.IssFee = 0
        If TCat$ = "ALL" Then
          Print #RptHandle, TCat$; dlm; TownRec.IssFee
        Else
          Print #RptHandle, TCat$ + "/" + GetCatDesc(TCat$); dlm; TownRec.IssFee
        End If
        CustCnt = CustCnt + 1
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustIdxRecs
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
Inactive:
  Next cnt
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  Close         'Close all open files now
  
  If CustCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers on file that fit the criteria entered."
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLCustListRpt.Show
    frmBLLoadReport.Show
  End If
  
  
  MainLog ("The 'Customer Detail Report' was processed in graphics format.")
  Exit Sub
  
  
GetCustFee:
  CustFee# = 0
  FeeAmt# = 0
  Prorate# = CustRec.Prorate
  If Prorate# >= 100 Or Prorate# < 0 Then
    Prorate# = 1
  Else
    Prorate# = OldRound(Prorate# * 0.01)
  End If
  
   CatCode$ = QPTrim$(CustRec.BILLCAT1)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          CustRec.DESC1 = CodeRec.CODEDESC           'Reset Code Descriptions
          If CodeRec.CodeType = "F" Then
            FeeAmt# = OldRound(Prorate# * CodeRec.Fee)
            GoTo C2
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV1
            FeeAmt# = OldRound#(Mult * CodeRec.Fee)
            FeeAmt# = OldRound(Prorate# * FeeAmt#)
            GoTo C2
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV1
            If Revenue# <= CodeRec.Recpt1 Then
              FeeAmt# = CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1)
              If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              FeeAmt# = CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2)
              If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              FeeAmt# = CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3)
              If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              FeeAmt# = CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4)
              If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              FeeAmt# = CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5)
              If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              FeeAmt# = CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6)
              If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C2
            End If
          End If '= "S"
        End If    'QPTrim$(CodeRec.CatCode) = CatCode$
      Next Snt&
    End If        'Len(CatCode$) > 0
    
    
C2:               'Catagory #2
    CustFee# = OldRound#(CustFee# + FeeAmt#)
    FeeAmt# = 0
    CatCode$ = QPTrim$(CustRec.BILLCAT2)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          If CodeRec.CodeType = "F" Then
            FeeAmt# = OldRound(Prorate# * CodeRec.Fee)
            GoTo C3
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV2
            FeeAmt# = OldRound#(Mult * CodeRec.Fee)
            FeeAmt# = OldRound(Prorate# * FeeAmt#)
            GoTo C3
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV2
            If Revenue# <= CodeRec.Recpt1 Then
              FeeAmt# = CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1)
              If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              FeeAmt# = CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2)
              If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              FeeAmt# = CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3)
              If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              FeeAmt# = CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4)
              If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              FeeAmt# = CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5)
              If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              FeeAmt# = CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6)
              If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo C3
            End If
          End If
        End If    'End Test for Code
      Next Snt&
    End If        'End Test for Cat 1
    
C3:
    CustFee# = OldRound#(CustFee# + FeeAmt#)
    FeeAmt# = 0
    CatCode$ = QPTrim$(CustRec.BILLCAT3)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          If CodeRec.CodeType = "F" Then
            FeeAmt# = OldRound(Prorate# * CodeRec.Fee)
            GoTo c4
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV3
            FeeAmt# = OldRound#(Mult * CodeRec.Fee)
            FeeAmt# = OldRound(Prorate# * FeeAmt#)
            GoTo c4
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV3
            If Revenue# <= CodeRec.Recpt1 Then
              FeeAmt# = CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1)
              If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              FeeAmt# = CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2)
              If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              FeeAmt# = CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3)
              If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              FeeAmt# = CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4)
              If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              FeeAmt# = CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5)
              If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              FeeAmt# = CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6)
              If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c4
            End If
          End If
        End If    'End Test for Code
      Next Snt&
    End If        'End Test for Cat 3
    
c4:
    CustFee# = OldRound#(CustFee# + FeeAmt#)
    FeeAmt# = 0
    CatCode$ = QPTrim$(CustRec.BILLCAT4)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          If CodeRec.CodeType = "F" Then
            FeeAmt# = OldRound(Prorate# * CodeRec.Fee)
            GoTo c5
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV4
            FeeAmt# = OldRound#(Mult * CodeRec.Fee)
            FeeAmt# = OldRound(Prorate# * FeeAmt#)
            GoTo c5
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV4
            If Revenue# <= CodeRec.Recpt1 Then
              FeeAmt# = CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1)
              If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              FeeAmt# = CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2)
              If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              FeeAmt# = CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3)
              If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              FeeAmt# = CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4)
              If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              FeeAmt# = CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5)
              If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              FeeAmt# = CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6)
              If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo c5
            End If
            
          End If
        End If    'End Test for Code
      Next Snt&
    End If        'End Test for Cat 1
    
c5:
    CustFee# = OldRound#(CustFee# + FeeAmt#)
    FeeAmt# = 0
    CatCode$ = QPTrim$(CustRec.BILLCAT5)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          If CodeRec.CodeType = "F" Then
            FeeAmt# = OldRound(Prorate# * CodeRec.Fee)
            GoTo SkipEm
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV5
            FeeAmt# = OldRound#(Mult * CodeRec.Fee)
            FeeAmt# = OldRound(Prorate# * FeeAmt#)
            GoTo SkipEm
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV5
            If Revenue# <= CodeRec.Recpt1 Then
              FeeAmt# = CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1)
              If FeeAmt# < CodeRec.BaseAmt1 Then FeeAmt# = CodeRec.BaseAmt1
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              FeeAmt# = CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2)
              If FeeAmt# < CodeRec.BaseAmt2 Then FeeAmt# = CodeRec.BaseAmt2
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              FeeAmt# = CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3)
              If FeeAmt# < CodeRec.BaseAmt3 Then FeeAmt# = CodeRec.BaseAmt3
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              FeeAmt# = CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4)
              If FeeAmt# < CodeRec.BaseAmt4 Then FeeAmt# = CodeRec.BaseAmt4
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              FeeAmt# = CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5)
              If FeeAmt# < CodeRec.BaseAmt5 Then FeeAmt# = CodeRec.BaseAmt5
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              FeeAmt# = CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6)
              If FeeAmt# < CodeRec.BaseAmt6 Then FeeAmt# = CodeRec.BaseAmt6
              FeeAmt# = OldRound(Prorate# * FeeAmt#)
              GoTo SkipEm
            End If
          End If
        End If
      Next Snt&
    End If
SkipEm:
  CustFee# = OldRound#(CustFee# + FeeAmt# + TownRec.IssFee)
  FeeAmt# = 0
  
  Return
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustListRpt", "PrintGraphics", Erl)
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
  If Not IsNumeric(fptxtCatCode.Text) Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub
