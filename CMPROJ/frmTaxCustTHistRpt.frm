VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxCustTHistRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Transaction History"
   ClientHeight    =   8736
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11640
   Icon            =   "frmTaxCustTHistRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8736
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   1680
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7470
      Left            =   1920
      TabIndex        =   5
      Top             =   630
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   13176
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.81
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
      Picture         =   "frmTaxCustTHistRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbDataType 
         Height          =   384
         Left            =   2928
         TabIndex        =   3
         Top             =   5040
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
         ColDesigner     =   "frmTaxCustTHistRpt.frx":08E6
      End
      Begin LpLib.fpList fpList 
         Height          =   912
         Left            =   1320
         TabIndex        =   2
         Top             =   3360
         Width           =   5292
         _Version        =   196608
         _ExtentX        =   9334
         _ExtentY        =   1609
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Columns         =   4
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   1
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
         ColDesigner     =   "frmTaxCustTHistRpt.frx":0D31
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   4
         Top             =   5688
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
         ColDesigner     =   "frmTaxCustTHistRpt.frx":11C1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   3156
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   6576
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":160C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5352
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmTaxCustTHistRpt.frx":17EA
         Top             =   6576
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":1895
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdMessage 
         Height          =   636
         Left            =   960
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6576
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":1A74
      End
      Begin EditLib.fpLongInteger fptxtAcctNum 
         Height          =   390
         Left            =   3405
         TabIndex        =   0
         Top             =   1680
         Width           =   1335
         _Version        =   196608
         _ExtentX        =   2355
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
         AlignTextH      =   1
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
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   661
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":1C52
      End
      Begin EditLib.fpText fptxtName 
         Height          =   390
         Left            =   1320
         TabIndex        =   1
         Top             =   2520
         Width           =   5295
         _Version        =   196608
         _ExtentX        =   9340
         _ExtentY        =   688
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
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
      Begin fpBtnAtlLibCtl.fpBtn cmdPropDet 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4440
         Width           =   2175
         _Version        =   131072
         _ExtentX        =   3836
         _ExtentY        =   661
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":1E6C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAll 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
         _ExtentY        =   661
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":208A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1215
         _Version        =   131072
         _ExtentX        =   2143
         _ExtentY        =   661
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
         ButtonDesigner  =   "frmTaxCustTHistRpt.frx":2267
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Property Listing:"
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
         Left            =   3000
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Data Type:"
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
         TabIndex        =   15
         Top             =   5145
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number:"
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
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
         TabIndex        =   9
         Top             =   5790
         Width           =   1500
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1050
         Top             =   315
         Width           =   5865
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Customer Transaction History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   5175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   765
         Top             =   1365
         Width           =   6330
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7740
      Left            =   1800
      Top             =   495
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxCustTHistRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Public RealRec As Long
  Public PersRec As Long
  Dim Town$
  Dim BtnFnt#
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim TempAcctNum As Long
  Dim ExitOK As Boolean

Private Sub cmdAll_Click()
  If fpList.ListCount > 0 Then
    fpList.Action = ActionSelectAll
  End If
End Sub

Private Sub cmdClear_Click()
  If fpList.ListCount > 0 Then
    fpList.Action = ActionDeselectAll
  End If
End Sub

Private Sub cmdExit_Click()
  ExitOK = True
  KillFile "custtranshist.dat"
  TempAcctNum = 0
  'frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLookup_Click()
  frmTaxCustLookup.Show
  DoEvents
End Sub

Private Sub cmdMessage_Click()
  If GCustNum > 0 Then
    frmTaxMessage.Show vbModal
  End If

End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbDataType.Text, "All") Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphics
    Else
      frmTaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      Call PrintText
    End If
  Else
    If fpList.ListCount > 1 Then
      If fpList.SelCount = 0 Then
        Call TaxMsg(900, "Please select a property on which to report.")
        Exit Sub
      End If
    End If
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsByProp
    Else
      frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      Call PrintTextByProp
    End If
  End If
  
End Sub

Private Sub cmdPropDet_Click()
  Dim ThisClass$
  Dim x As Integer
  
  If fpList.SelCount = 0 Then
    Call TaxMsg(900, "Please make a selection from the list.")
    Exit Sub
  ElseIf fpList.SelCount > 1 Then
    Call TaxMsg(900, "Please clear the list and reselect a property.")
    Exit Sub
  End If
  
  For x = 0 To fpList.ListCount - 1
    fpList.Row = x
    If fpList.Selected = True Then
      fpList.ListIndex = x
    End If
  Next x
  fpList.Row = fpList.ListIndex
  fpList.col = 0
  ThisClass = QPTrim$(fpList.ColText)
  fpList.col = 3
  If ThisClass = "PERSONAL" Then
    PersRec = CLng(fpList.ColText)
    frmTaxPersDetail.Show vbModal
    Exit Sub
  ElseIf ThisClass = "REAL" Then
    RealRec = CLng(fpList.ColText)
    frmTaxRealDetail.Show vbModal
    Exit Sub
  Else
    Call TaxMsg(800, "The classification for the selected property cannot be determined. Detail data cannot be loaded.")
    Exit Sub
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
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%M"
      Call cmdMessage_Click
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
  Call Loadme
  ExitOK = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "custtranshist.dat"
      TXLog ("CitiTaxes.exe terminated via menu bar on frmTaxCustTransHist.")
      Call CMTerminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Loadme()
  Dim One As Integer
  Dim AHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbDataType.Text = "List All Transactions"
  fpcmbDataType.AddItem "List All Transactions"
  fpcmbDataType.AddItem "List Transactions By Property"
  
  One = 1
  AHandle = FreeFile
  Open "custtranshist.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
End Sub

Private Sub fpcmbDataType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDataType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDataType.ListIndex = -1
  End If
  If fpcmbDataType.ListDown <> True Then
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
      fptxtAcctNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxType_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbTaxType.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbTaxType.ListIndex = -1
'  End If
'  If fpcmbTaxType.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fptxtBegDate.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
End Sub

Private Sub fpcmbTransType_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbTransType.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbTransType.ListIndex = -1
'  End If
'  If fpcmbTransType.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fpcmbTaxType.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If

End Sub

Private Sub fptxtAcctNum_LostFocus()
  If ExitOK = True Then
    ExitOK = False
    Exit Sub
  End If
  If CLng(fptxtAcctNum.Text) = 0 Then Exit Sub
  If TempAcctNum = CLng(fptxtAcctNum.Text) Then Exit Sub
  If Check4ValidCustNum(CLng(fptxtAcctNum.Text)) = True Then
    GCustNum = CLng(fptxtAcctNum.Value)
    Call LoadCust
  Else
    Call TaxMsg(900, "The account number entered is not valid.")
    Call Clearscreen
    fptxtAcctNum.Text = "0"
    fptxtAcctNum.SetFocus
    TempAcctNum = 0
  End If

End Sub

Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdMessage.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdMessage.ForeColor = &H80000012
      cmdMessage.FontSize = BtnFnt
    Case 2
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 0.7
    Case 3
      cmdMessage.ForeColor = &H80000011
      cmdMessage.FontSize = BtnFnt - 1.4
    Case 4
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.1
    Case 5
      cmdMessage.ForeColor = &H80000010
      cmdMessage.FontSize = BtnFnt - 2.8
    Case 6
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 3.5
    Case 7
      cmdMessage.ForeColor = &H8000000F
      cmdMessage.FontSize = BtnFnt - 4.2
    Case 8
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 4.9
    Case 9
      cmdMessage.ForeColor = &H8000000E
      cmdMessage.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
'  DoEvents
End Sub

Public Sub LoadCust()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  fptxtName.Text = QPTrim$(TaxCust.CustName)
  fptxtAcctNum = GCustNum
  TempAcctNum = GCustNum
  If GCustNum > 0 Then
    If CustHasMsg(GCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdMessage.ForeColor = &H80000012
    End If
  End If
  
  fpList.Clear
  OpenRealPropFile RHandle, NumOfRealRecs
  
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealPropRec
    If TaxCust.PIN = RealPropRec.CustPin Then
      fpList.InsertRow = "REAL" + Chr(9) + QPTrim$(RealPropRec.RealPin) + Chr(9) + QPTrim$(RealPropRec.RealPin) + Chr(9) + CStr(x)
    End If
  Next x
  Close RHandle

  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersPropRec
    If TaxCust.PIN = PersPropRec.CustPin Then
      fpList.InsertRow = "PERSONAL" + Chr(9) + QPTrim$(PersPropRec.PropPin) + Chr(9) + QPTrim$(PersPropRec.PropPin) + Chr(9) + CStr(x)
    End If
  Next x
  Close PHandle

End Sub
Private Function Check4ValidCustNum(ThisCust As Long) As Boolean
  Dim TaxRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim Number$
  Dim Name$
  Dim Found As Boolean

  Check4ValidCustNum = True
  
  If fptxtAcctNum.Value = 0 Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmTaxMsg.Label1.Caption = "There are no tax customers saved."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If ThisCust = TaxRec.Acct Then
      If TaxRec.Deleted <> 0 Then
        Check4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > NumOfCRecs Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
  
End Function

Private Sub Clearscreen()
  fptxtName.Text = ""
  TempAcctNum = 0
'  fptxtAcctNum = 0
End Sub

Private Sub PrintGraphics()
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim TotBal As Double
  
  If Check4ValidCustNum(GCustNum) = False Then
    Exit Sub
  End If
  
  TotBal = GetCustBalance(GCustNum, -1)
  dlm$ = "~"
    
  RptFile$ = "TAXRPTS\CHSTTRAN.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  ReDim TotByYrAndType(1 To 14, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 14, 1 To 1) As Integer
  ReDim TheYear(1 To 1) As Integer
  
  EnableCloseButton Me.hwnd, False
  Get TCHandle, GCustNum, TaxCust
  ThisName = QPTrim$(TaxCust.CustName)
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If YrCnt = 0 Then
       YrCnt = YrCnt + 1
       ThisYear = YrCnt
       ReDim Preserve TheYear(1 To YrCnt) As Integer
       TheYear(YrCnt) = TaxTrans.TaxYear
       ReDim Preserve TotByYrAndType(1 To 14, 1 To YrCnt) As Double
       ReDim Preserve CntByYrAndType(1 To 14, 1 To YrCnt) As Integer
       For y = 1 To 14
         TotByYrAndType(y, YrCnt) = 0
         CntByYrAndType(y, YrCnt) = 0
       Next y
     Else
       For y = 1 To YrCnt
         If TaxTrans.TaxYear = TheYear(y) Then
           ThisYear = y
           Exit For
         End If
       Next y
       If y > YrCnt Then
         YrCnt = YrCnt + 1
         ThisYear = YrCnt
         ReDim Preserve TheYear(1 To YrCnt) As Integer
         TheYear(YrCnt) = TaxTrans.TaxYear
         ReDim Preserve TotByYrAndType(1 To 14, 1 To YrCnt) As Double
         ReDim Preserve CntByYrAndType(1 To 14, 1 To YrCnt) As Integer
         For y = 1 To 14
           TotByYrAndType(y, YrCnt) = 0
           CntByYrAndType(y, YrCnt) = 0
         Next y
       End If
     End If
         
     Select Case TaxTrans.TranType
       Case 1
         ThisTransType = "Billing"
         TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
       Case 2
         ThisTransType = "Payment"
         TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
       Case 3
         ThisTransType = "Release"
         TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
       Case 4
         ThisTransType = "Interest"
         TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
       Case 5
         ThisTransType = "Unknown"
         TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
       Case 6
          ThisTransType = "Advertising Charge"
          TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
       Case 7
         ThisTransType = "Adjust Pay Down"
         TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
       Case 9
         ThisTransType = "Credit Applied at Billing"
         TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
       Case 13
         ThisTransType = "Adjust Bill Down"
         TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
       Case 14
         ThisTransType = "Adjust Bill Up"
         TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
       Case 21
         ThisTransType = "Billpay/Overpay"
         TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
       Case 22
         ThisTransType = "Overpayment"
         TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
       Case 24
         ThisTransType = "Adjust Bill Up Affecting Credit Balance"
         TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
       Case Else
         ThisTransType = "Unknown"
         TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
         CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
      End Select
      TCnt = TCnt + 1
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      '                   0            1                 2                   3
      Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
      '                                 4                           5                6
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
      '                          7                         8                          9
      Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
      '                      10                11          12                       13
      Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;

      If TaxTrans.BelongTo > 0 Then
        Get TTHandle, TaxTrans.BelongTo, TaxTrans
        '                             14
        Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
      Else
        '                 14
        Print #RptHandle, 0; dlm;
      End If
      Get TTHandle, ThisRec, TaxTrans
      '                                15                       16                17
      Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm; TotBal
SkipIt:
      ThisRec = TaxTrans.LastTrans
  Loop
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then GoSub PrintSub
  arTaxCustTHistByTrans.Show
  
  Exit Sub
  
PrintSub:
  
  SubRptFile$ = "TAXRPTS\SUBTAXJRNL.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  BigYr = 0
  For x = 1 To YrCnt
    If TheYear(x) > BigYr Then
      BigYr = TheYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 14, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 14, 1 To YrCnt) As Double
  
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
    For x = 1 To 14
      For y = Nexty To YrCnt
        If TheYear(y) >= HoldBigYr Then
          HoldBigYr = TheYear(y)
          Thisx = x
          Thisy = y
        End If
      Next y
      For z = 1 To 14
        HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
        HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
      Next z
      HoldYr = TheYear(Nexty)
      For z = 1 To 14
        TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
        CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
      Next z
      TheYear(Nexty) = TheYear(Thisy)
      For z = 1 To 14
        TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
        CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
      Next z
      TheYear(Thisy) = HoldYr
      If Nexty >= YrCnt Then Exit For
      HoldBigYr = 0
      Nexty = Nexty + 1
    Next x
  
  For y = 1 To YrCnt
    For x = 1 To 14
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #SubRptHandle, "Billing"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 2
            Print #SubRptHandle, "Payment"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 3
            Print #SubRptHandle, "Release"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 4
            Print #SubRptHandle, "Interest"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 5
            Print #SubRptHandle, "Unknown"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 6
            Print #SubRptHandle, "Advertising"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 7
            Print #SubRptHandle, "Adjust Pay Down"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 8
            Print #SubRptHandle, "Credit at Billing"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 9
            Print #SubRptHandle, "Adjust Bill Down"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 10
            Print #SubRptHandle, "Adjust Bill Up"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 11
            Print #SubRptHandle, "Bill OverPay"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 12
            Print #SubRptHandle, "OverPayment"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 13
            Print #SubRptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
          Case 14
            Print #SubRptHandle, "Unknown"; dlm; TheYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y)
        End Select
      End If
    Next x
  Next y
  Close SubRptHandle
  Return
  
End Sub

Private Sub PrintText()
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As TXCustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long, NewName$
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, Page As Integer
  Dim CustName$, PrintCnt As Integer
  Dim ThisBillNum As String * 8
  Dim TotBal As Double
  
  TotBal = GetCustBalance(GCustNum, -1)
  CustName = ""
  IdxFlag = False
  If Check4ValidCustNum(GCustNum) = False Then
    Exit Sub
  End If
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0

  RptFile$ = "TAXRPTS\TAXJRNL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  ReDim TotByYrAndType(1 To 14, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 14, 1 To 1) As Double
  ReDim TheYear(1 To 1) As Integer
    
  ThisName = QPTrim$(TaxCust.CustName)
  ThisRec = TaxCust.LastTrans
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  PrintCnt = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
      If PrintCnt = 0 Then
        If LineCnt <> 6 Then
          Print #RptHandle,
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
        End If
        GoSub PrintCustHeader
      End If
      PrintCnt = PrintCnt + 1
      If YrCnt = 0 Then
         YrCnt = YrCnt + 1
         ThisYear = YrCnt
         ReDim Preserve TheYear(1 To YrCnt) As Integer
         TheYear(YrCnt) = TaxTrans.TaxYear
         ReDim Preserve TotByYrAndType(1 To 14, 1 To YrCnt) As Double
         ReDim Preserve CntByYrAndType(1 To 14, 1 To YrCnt) As Double
         For y = 1 To 14
           TotByYrAndType(y, YrCnt) = 0
         Next y
       Else
         For y = 1 To YrCnt
           If TaxTrans.TaxYear = TheYear(y) Then
             ThisYear = y
             Exit For
           End If
         Next y
         If y > YrCnt Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve TheYear(1 To YrCnt) As Integer
           TheYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 14, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 14, 1 To YrCnt) As Double
           For y = 1 To 14
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
         End If
       End If
          
       Select Case TaxTrans.TranType
         Case 1
           ThisTransType = "Billing"
           TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
         Case 2
           ThisTransType = "Payment"
           TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
         Case 3
           ThisTransType = "Release"
           TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
         Case 4
           ThisTransType = "Interest"
           TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
         Case 5
           ThisTransType = "Unknown"
           TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
         Case 6
           ThisTransType = "Advertising Charge"
           TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
         Case 7
           ThisTransType = "Adjust Pay Down"
           TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
         Case 9
           ThisTransType = "Credit Applied at Billing"
           TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
         Case 13
           ThisTransType = "Adjust Bill Down"
           TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
         Case 14
           ThisTransType = "Adjust Bill Up"
           TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
         Case 21
           ThisTransType = "Billpay/Overpay"
           TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
         Case 22
           ThisTransType = "Overpayment"
           TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
         Case 24
           ThisTransType = "Adjust Bill Up Affecting Credit Balance"
           TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
         Case Else
           ThisTransType = "Unknown"
           TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
           CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
      End Select
      TCnt = TCnt + 1
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
      Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
      Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(69);
      LineCnt = LineCnt + 1
      If TaxTrans.BelongTo > 0 Then
        Get TTHandle, TaxTrans.BelongTo, TaxTrans
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        If IsNumeric(ThisBillNum) Then
          Print #RptHandle, Using$("######", CDbl(ThisBillNum));
        Else
          Print #RptHandle, "   " + ThisBillNum;
        End If
      Else
        Print #RptHandle, "     0";
      End If
      
      Get TTHandle, ThisRec, TaxTrans
      Print #RptHandle, Tab(79); ThisTransType
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
SkipIt:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  If YrCnt > 0 Then GoSub SortIt
  Print #RptHandle, FF$
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  ViewPrint RptFile, "Tax Transactions Report", True
  
  Exit Sub
  
SortIt:
  
  BigYr = 0
  For x = 1 To YrCnt
    If TheYear(x) > BigYr Then
      BigYr = TheYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 14, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 14, 1 To YrCnt) As Double
  
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 14
    For y = Nexty To YrCnt
      If TheYear(y) >= HoldBigYr Then
        HoldBigYr = TheYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 14
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = TheYear(Nexty)
    For z = 1 To 14
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    TheYear(Nexty) = TheYear(Thisy)
    For z = 1 To 14
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    TheYear(Thisy) = HoldYr
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0 'BigYr + 1
    Nexty = Nexty + 1
  Next x
  Print #RptHandle, FF$
  GoSub PrintSortHeader
  LineCnt = LineCnt + 3
  For y = 1 To YrCnt
   If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintSortHeader
      Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
    End If
    Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
    LineCnt = LineCnt + 1
    For x = 1 To 13
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #RptHandle, "  Billing"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 2
            Print #RptHandle, "  Payment"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 3
            Print #RptHandle, "  Release"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 4
            Print #RptHandle, "  Interest"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y))  'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 5
            Print #RptHandle, "  Unknown"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 6
            Print #RptHandle, "  Advertising"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 7
            Print #RptHandle, "  Adjust Pay Down"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 8
            Print #RptHandle, "  Credit at Billing"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 9
            Print #RptHandle, "  Adjust Bill Down"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 10
            Print #RptHandle, "  Adjust Bill Up"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 11
            Print #RptHandle, "  Bill OverPay"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 12
            Print #RptHandle, "  OverPayment"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 13
            Print #RptHandle, "  Adj Bill Up Affecting Credit"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
          Case 14
            Print #RptHandle, "  Unknown"; Tab(30); Using$("##,##0", CntByYrAndType(x, y)); Tab(50); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", TheYear(y))
            End If
        End Select
      End If
    Next x
    Print #RptHandle, String$(94, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 2
  Next y
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Customer Transaction History"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle,
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(44); "Overpay Amt"; Tab(57); "Trans Amt"; Tab(67); "Belongs To"; Tab(78); "Trans Type"
  Print #RptHandle, String(94, "-")
  LineCnt = 6
  
  Return
  
PrintCustHeader:
  If LineCnt <> 6 Then
    Print #RptHandle, String(94, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Cust Num: " + Using$("#######0", TaxCust.Acct); Tab(21); "Customer Name: "; Tab(37); QPTrim$(TaxCust.CustName); Tab(80); "Active: "; Tab(89); TaxCust.Active
  Print #RptHandle, "Total Balance: " + QPTrim$(Using$("$###,###,##0.00", TotBal))
  Print #RptHandle, String(94, ".")
  LineCnt = LineCnt + 3
  
  Return
  
PrintSortHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Customer Transaction History Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Total Transaction Count: " + Using$("#####0", TCnt)
  Print #RptHandle, "Description"; Tab(30); "Trans Count"; Tab(59); "Amount"
  Print #RptHandle, String$(94, "-")
  LineCnt = 6
  
  Return

GetTransType:
  Select Case TaxTrans.TranType
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Overpayment"
    Case Else
      ThisType = "All"
  End Select
 
  Return
End Sub

Private Sub PrintGraphicsByProp()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer, z As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim ThisTransType As String
  Dim PropCnt As Integer
  Dim RealPropPin As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersPropPin As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim PropTransCnt As Integer
  Dim BillRec As Long
  Dim CustPin As Long
  Dim ThisBillNum$
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim BalThisBill As Double
  Dim TotBal As Double
  Dim TotPropCnt As Integer
  
  If Check4ValidCustNum(GCustNum) = False Then
    Exit Sub
  End If
  
  dlm$ = "~"
  PropCnt = 0
  ReDim PropPin(1 To 1) As String
  ReDim PropType(1 To 1) As String
  If fpList.ListCount = 1 Then
    fpList.Row = fpList.ListIndex
    fpList.col = 2
    PropCnt = PropCnt + 1
    ReDim Preserve PropPin(1 To PropCnt) As String
    PropPin(PropCnt) = QPTrim(fpList.ColText)
    fpList.col = 0
    ReDim Preserve PropType(1 To PropCnt) As String
    PropType(PropCnt) = Mid(fpList.ColText, 1, 1)
  Else
    For x = 0 To fpList.ListCount - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        fpList.col = 2
        PropCnt = PropCnt + 1
        ReDim Preserve PropPin(1 To PropCnt) As String
        PropPin(PropCnt) = QPTrim(fpList.ColText)
        fpList.col = 0
        ReDim Preserve PropType(1 To PropCnt) As String
        PropType(PropCnt) = Mid(fpList.ColText, 1, 1)
      End If
     Next x
  End If
    
  RptFile$ = "TAXRPTS\CHSTPROP.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  CustPin = TaxCust.PIN
  TotBal = GetCustBalance(GCustNum, -1)
  OpenTaxTransFile TTHandle, NumOfTTRecs
  ThisName = QPTrim$(TaxCust.CustName)
  ThisRec = TaxCust.LastTrans
  
  FrmShowPctComp.Label1 = "Gathering Tax Transactions"
  FrmShowPctComp.Show
'  frmShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  For y = 1 To PropCnt
    PropTransCnt = 0
    ReDim ThisPropTrans(1 To 1) As Long
    For x = 1 To NumOfTTRecs
      Get TTHandle, x, TaxTrans
      If PropType(y) = "P" Then
        If TaxTrans.PersPin = PropPin(y) And TaxTrans.CustPin = CustPin Then
          PropTransCnt = PropTransCnt + 1
          ReDim Preserve ThisPropTrans(1 To PropTransCnt) As Long
          ThisPropTrans(PropTransCnt) = x
        End If
      End If
      FrmShowPctComp.ShowPctComp x, NumOfTTRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
    Next x
    For x = 1 To PropTransCnt
      Get TTHandle, ThisPropTrans(x), TaxTrans
      If TaxTrans.TranType = 1 Then
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        BalThisBill = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        BillRec = ThisPropTrans(x)
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        GoSub GetType
        TotPropCnt = TotPropCnt + 1
        '                   0             1               2                        3
        Print #RptHandle, Town$; dlm; "PERSONAL"; dlm; ThisType; dlm; QPTrim$(TaxTrans.PersPin); dlm;
        '                               4                           5                          6
        Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; TaxTrans.Amount; dlm; MakeRegDate(TaxTrans.TransDate); dlm;
        '                          7                 8                 9              10            11                        12
        Print #RptHandle, TaxTrans.TaxYear; dlm; ThisBillNum; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TaxTrans.Revenue.Principle1; dlm;
        '                            13                               14                               15
        Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.LateList; dlm;
        '                            16                               17                            18
        Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt3; dlm;
        '                            19                               20                                   21
        Print #RptHandle, TaxTrans.Revenue.Principle1Pd; dlm; TaxTrans.Revenue.InterestPd; dlm; TaxTrans.Revenue.CollectionPd; dlm;
        '                            22                               23                                   24
        Print #RptHandle, TaxTrans.Revenue.LateListPd; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm;
        '                            25                       26             24           28            29
        Print #RptHandle, TaxTrans.Revenue.RevOpt3Pd; dlm; PrincDif; dlm; IntDif; dlm; AdvDif; dlm; LateListDif; dlm;
        '                    30            31            32                33                   34          35
        Print #RptHandle, Opt1Dif; dlm; Opt2Dif; dlm; Opt3Dif; dlm; TaxTrans.TranType; dlm; BalThisBill; dlm; TotBal; dlm;
        '                     36                 37
        Print #RptHandle, TaxCust.PIN; dlm; QPTrim$(TaxCust.CustName)
        
        For z = 1 To PropTransCnt
          Get TTHandle, ThisPropTrans(z), TaxTrans
          If TaxTrans.BelongTo = BillRec Then
            PrincDif = 0
            IntDif = 0
            AdvDif = 0
            LateListDif = 0
            Opt1Dif = 0
            Opt2Dif = 0
            Opt3Dif = 0
            GoSub GetType
            TotPropCnt = TotPropCnt + 1
            '                   0              1              2                          3
            Print #RptHandle, Town$; dlm; "PERSONAL"; dlm; ThisType; dlm; QPTrim$(TaxTrans.PersPin); dlm;
            '                            4                              5                           6
            Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; TaxTrans.Amount; dlm; MakeRegDate(TaxTrans.TransDate); dlm;
            '                          7                 8                 9              10            11                        12
            Print #RptHandle, TaxTrans.TaxYear; dlm; ThisBillNum; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TaxTrans.Revenue.Principle1; dlm;
            '                            13                               14                               15
            Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.LateList; dlm;
            '                            16                               17                            18
            Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt3; dlm;
            '                            19                               20                                   21
            Print #RptHandle, TaxTrans.Revenue.Principle1Pd; dlm; TaxTrans.Revenue.InterestPd; dlm; TaxTrans.Revenue.CollectionPd; dlm;
            '                            22                               23                                   24
            Print #RptHandle, TaxTrans.Revenue.LateListPd; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm;
            '                            25                       26             24           28            29
            Print #RptHandle, TaxTrans.Revenue.RevOpt3Pd; dlm; PrincDif; dlm; IntDif; dlm; AdvDif; dlm; LateListDif; dlm;
            '                    30            31            32               33                34        35
            Print #RptHandle, Opt1Dif; dlm; Opt2Dif; dlm; Opt3Dif; dlm; TaxTrans.TranType; dlm; 0; dlm; TotBal; dlm;
            '                    36                     37
            Print #RptHandle, TaxCust.PIN; dlm; QPTrim$(TaxCust.CustName)
          End If
        Next z
      End If
    Next x
  Next y
  
  FrmShowPctComp.Label1 = "Gathering Tax Transactions"
  FrmShowPctComp.Show
  For y = 1 To PropCnt
    PropTransCnt = 0
    ReDim ThisPropTrans(1 To 1) As Long
    For x = 1 To NumOfTTRecs
'      If x = 19393 Then Stop
      Get TTHandle, x, TaxTrans
      If PropType(y) = "R" Then
        If QPTrim$(TaxTrans.RealPin) = PropPin(y) And TaxTrans.CustPin = CustPin Then
          PropTransCnt = PropTransCnt + 1
          ReDim Preserve ThisPropTrans(1 To PropTransCnt) As Long
          ThisPropTrans(PropTransCnt) = x
        End If
      End If
      FrmShowPctComp.ShowPctComp x, NumOfTTRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
    Next x

    For x = 1 To PropTransCnt
      Get TTHandle, ThisPropTrans(x), TaxTrans
      If TaxTrans.TranType = 1 Then
        BillRec = ThisPropTrans(x)
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        BalThisBill = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        GoSub GetType
        TotPropCnt = TotPropCnt + 1
        '                   0           1             2                      3
        Print #RptHandle, Town; dlm; "REAL"; dlm; ThisType; dlm; QPTrim(TaxTrans.RealPin); dlm;
        '                               4                            5                           6
        Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; TaxTrans.Amount; dlm; MakeRegDate(TaxTrans.TransDate); dlm;
        '                          7                 8                 9              10            11                        12
        Print #RptHandle, TaxTrans.TaxYear; dlm; ThisBillNum; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TaxTrans.Revenue.Principle1; dlm;
        '                            13                               14                               15
        Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.LateList; dlm;
        '                            16                               17                            18
        Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt3; dlm;
        '                            19                               20                                   21
        Print #RptHandle, TaxTrans.Revenue.Principle1Pd; dlm; TaxTrans.Revenue.InterestPd; dlm; TaxTrans.Revenue.CollectionPd; dlm;
        '                            22                               23                                   24
        Print #RptHandle, TaxTrans.Revenue.LateListPd; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm;
        '                            25                       26             24           28            29
        Print #RptHandle, TaxTrans.Revenue.RevOpt3Pd; dlm; PrincDif; dlm; IntDif; dlm; AdvDif; dlm; LateListDif; dlm;
        '                    30            31            32                 33                  34              35
        Print #RptHandle, Opt1Dif; dlm; Opt2Dif; dlm; Opt3Dif; dlm; TaxTrans.TranType; dlm; BalThisBill; dlm; TotBal; dlm;
        '                     36                      37
        Print #RptHandle, TaxCust.PIN; dlm; QPTrim$(TaxCust.CustName)
        For z = 1 To PropTransCnt
          Get TTHandle, ThisPropTrans(z), TaxTrans
          If TaxTrans.BelongTo = BillRec Then
            PrincDif = 0
            IntDif = 0
            AdvDif = 0
            LateListDif = 0
            Opt1Dif = 0
            Opt2Dif = 0
            Opt3Dif = 0
            GoSub GetType
            TotPropCnt = TotPropCnt + 1
            '                   0          1              2                      3
            Print #RptHandle, Town; dlm; "REAL"; dlm; ThisType; dlm; QPTrim$(TaxTrans.RealPin); dlm;
            '                                4                          5                               6
            Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; TaxTrans.Amount; dlm; MakeRegDate(TaxTrans.TransDate); dlm;
            '                          7                 8                 9              10            11                        12
            Print #RptHandle, TaxTrans.TaxYear; dlm; ThisBillNum; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TaxTrans.Revenue.Principle1; dlm;
            '                            13                               14                               15
            Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.LateList; dlm;
            '                            16                               17                            18
            Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt3; dlm;
            '                            19                               20                                   21
            Print #RptHandle, TaxTrans.Revenue.Principle1Pd; dlm; TaxTrans.Revenue.InterestPd; dlm; TaxTrans.Revenue.CollectionPd; dlm;
            '                            22                               23                                   24
            Print #RptHandle, TaxTrans.Revenue.LateListPd; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm;
            '                            25                       26             24           28            29
            Print #RptHandle, TaxTrans.Revenue.RevOpt3Pd; dlm; PrincDif; dlm; IntDif; dlm; AdvDif; dlm; LateListDif; dlm;
            '                    30            31            32                33               34       35
            Print #RptHandle, Opt1Dif; dlm; Opt2Dif; dlm; Opt3Dif; dlm; TaxTrans.TranType; dlm; 0; dlm; TotBal; dlm;
            '                      36                     37
            Print #RptHandle, TaxCust.PIN; dlm; QPTrim$(TaxCust.CustName)
          End If
        Next z
      End If
    Next x
MoveOn:
  Next y
  Unload FrmShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  If TotPropCnt = 0 Then
    Call TaxMsg(800, "There are no transactions saved containing the pointers necessary to generate this report.")
    Exit Sub
  End If
  
  arTaxCustTHistByProp.Show
  Exit Sub
  
GetType:
  Select Case TaxTrans.TranType
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit Applied at Billing"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Overpayment"
    Case 22
      ThisType = "Paid Bill Plus Overpay"
'    Case 23
'      ThisType = "Adjust Bill Down Affecting Credit Balance"
    Case 24
      ThisType = "Adjust Bill Up Affecting Credit Balance"
    Case Else
      ThisType = "Unknown"
  End Select

  Return
End Sub

Private Sub PrintTextByProp()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer, z As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim ThisTransType As String
  Dim PropCnt As Integer
  Dim RealPropPin As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersPropPin As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim PropTransCnt As Integer
  Dim BillRec As Long
  Dim CustPin As Long
  Dim ThisBillNum$
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim BalThisBill As Double
  Dim TotBal As Double
  Dim TotPropCnt As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim FF$, ThisPin$, RealPers$
  
  FF$ = Chr(12)
  MaxLines = 58
  If Check4ValidCustNum(GCustNum) = False Then
    Exit Sub
  End If

  PropCnt = 0
  ReDim PropPin(1 To 1) As String
  ReDim PropType(1 To 1) As String
  If fpList.ListCount = 1 Then
    fpList.Row = fpList.ListIndex
    fpList.col = 2
    PropCnt = PropCnt + 1
    ReDim Preserve PropPin(1 To PropCnt) As String
    PropPin(PropCnt) = QPTrim(fpList.ColText)
    fpList.col = 0
    ReDim Preserve PropType(1 To PropCnt) As String
    PropType(PropCnt) = Mid(fpList.ColText, 1, 1)
  Else
    For x = 0 To fpList.ListCount - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        fpList.col = 2
        PropCnt = PropCnt + 1
        ReDim Preserve PropPin(1 To PropCnt) As String
        PropPin(PropCnt) = QPTrim(fpList.ColText)
        fpList.col = 0
        ReDim Preserve PropType(1 To PropCnt) As String
        PropType(PropCnt) = Mid(fpList.ColText, 1, 1)
      End If
     Next x
  End If
  
  RptFile$ = "TAXRPTS\CHSTPROP.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  CustPin = TaxCust.PIN
  TotBal = GetCustBalance(GCustNum, -1)
  OpenTaxTransFile TTHandle, NumOfTTRecs
  ThisName = QPTrim$(TaxCust.CustName)
  ThisRec = TaxCust.LastTrans

  GoSub PrintHeader

  FrmShowPctComp.Label1 = "Gathering Tax Transactions"
  FrmShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  For y = 1 To PropCnt
    PropTransCnt = 0
    ReDim ThisPropTrans(1 To 1) As Long
    For x = 1 To NumOfTTRecs
      Get TTHandle, x, TaxTrans
      If PropType(y) = "P" Then
        If TaxTrans.PersPin = PropPin(y) And TaxTrans.CustPin = CustPin Then
          PropTransCnt = PropTransCnt + 1
          ReDim Preserve ThisPropTrans(1 To PropTransCnt) As Long
          ThisPropTrans(PropTransCnt) = x
        End If
      End If
      FrmShowPctComp.ShowPctComp x, NumOfTTRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
    Next x
    RealPers = "PERSONAL"
    If PropTransCnt > 1 Then
      Print #RptHandle, RealPers
      Print #RptHandle, String(79, "-")
      LineCnt = LineCnt + 2
    End If
    For x = 1 To PropTransCnt
      Get TTHandle, ThisPropTrans(x), TaxTrans
      ThisPin = QPTrim$(TaxTrans.PersPin)
      If TaxTrans.TranType = 1 Then
        GoSub PrintBillHeader
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        BalThisBill = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        BillRec = ThisPropTrans(x)
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        GoSub GetType
        TotPropCnt = TotPropCnt + 1
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt = 1
        End If
        Print #RptHandle, Tab(20); "Principal    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(55); Using("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(70); Using$("$##,##0.00", PrincDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt = 1
        End If
        Print #RptHandle, Tab(20); "Interest     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(70); Using$("$##,##0.00", IntDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt = 1
        End If
        Print #RptHandle, Tab(20); "Advertising  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd); Tab(70); Using$("$##,##0.00", AdvDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt = 1
        End If
        Print #RptHandle, Tab(20); "Late Listing "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(70); Using$("$##,##0.00", LateListDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt = 1
        End If
        If Len(QPTrim$(Opt1Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(70); Using$("$##,##0.00", Opt1Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt = 1
          End If
        End If
        If Len(QPTrim$(Opt2Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(70); Using$("$##,##0.00", Opt2Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt = 1
          End If
        End If
        If Len(QPTrim$(Opt3Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(70); Using$("$##,##0.00", Opt3Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt = 1
          End If
        End If
        Print #RptHandle, String(79, "-")
        Print #RptHandle,
        LineCnt = LineCnt + 2
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        For z = 1 To PropTransCnt
          Get TTHandle, ThisPropTrans(z), TaxTrans
          If TaxTrans.BelongTo = BillRec Then
            PrincDif = 0
            IntDif = 0
            AdvDif = 0
            LateListDif = 0
            Opt1Dif = 0
            Opt2Dif = 0
            Opt3Dif = 0
            GoSub GetType
            TotPropCnt = TotPropCnt + 1
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Principal    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(55); Using("$##,##0.00", TaxTrans.Revenue.Principle1Pd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Interest     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Advertising  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Late Listing "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            If Len(QPTrim$(Opt1Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            If Len(QPTrim$(Opt2Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            If Len(QPTrim$(Opt3Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            Print #RptHandle, String(79, "-")
            Print #RptHandle,
            LineCnt = LineCnt + 2
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
        Next z
      End If
    Next x
  Next y

  FrmShowPctComp.Label1 = "Gathering Tax Transactions"
  FrmShowPctComp.Show
  For y = 1 To PropCnt
    PropTransCnt = 0
    ReDim ThisPropTrans(1 To 1) As Long
    For x = 1 To NumOfTTRecs
      Get TTHandle, x, TaxTrans
      If PropType(y) = "R" Then
        If QPTrim$(TaxTrans.RealPin) = PropPin(y) And TaxTrans.CustPin = CustPin Then
          PropTransCnt = PropTransCnt + 1
          ReDim Preserve ThisPropTrans(1 To PropTransCnt) As Long
          ThisPropTrans(PropTransCnt) = x
        End If
      End If
      FrmShowPctComp.ShowPctComp x, NumOfTTRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        EnableCloseButton Me.hwnd, True
        Exit Sub
      End If
    Next x
    RealPers = "REAL"
    If PropTransCnt > 1 Then
      Print #RptHandle, RealPers
      Print #RptHandle, String(79, "-")
      LineCnt = LineCnt + 2
    End If

    For x = 1 To PropTransCnt
      Get TTHandle, ThisPropTrans(x), TaxTrans
      ThisPin = QPTrim$(TaxTrans.RealPin)

      If TaxTrans.TranType = 1 Then
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        GoSub PrintBillHeader
        BillRec = ThisPropTrans(x)
        ThisBillNum = ParseBillNum(TaxTrans.Description)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        BalThisBill = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        GoSub GetType
        TotPropCnt = TotPropCnt + 1
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt + 1
        End If
        Print #RptHandle, Tab(20); "Principal    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(55); Using("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(70); Using$("$##,##0.00", PrincDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt + 1
        End If
        Print #RptHandle, Tab(20); "Interest     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(70); Using$("$##,##0.00", IntDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt + 1
        End If
        Print #RptHandle, Tab(20); "Advertising  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd); Tab(70); Using$("$##,##0.00", AdvDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt + 1
        End If
        Print #RptHandle, Tab(20); "Late Listing "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(70); Using$("$##,##0.00", LateListDif)
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, RealPers
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 2
          GoSub PrintBillHeader
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
          LineCnt = LineCnt + 1
        End If
        If Len(QPTrim$(Opt1Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(70); Using$("$##,##0.00", Opt1Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt + 1
          End If
        End If
        If Len(QPTrim$(Opt2Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(70); Using$("$##,##0.00", Opt2Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt + 1
          End If
        End If
        If Len(QPTrim$(Opt3Desc)) > 0 Then
          Print #RptHandle, Tab(20); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(70); Using$("$##,##0.00", Opt3Dif)
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
            Print #RptHandle, RealPers
            Print #RptHandle, String(79, "-")
            LineCnt = LineCnt + 2
            GoSub PrintBillHeader
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount); Tab(70); Using$("$##,##0.00", BalThisBill)
            LineCnt = LineCnt + 1
          End If
        End If
        Print #RptHandle, String(79, "-")
        Print #RptHandle,
        LineCnt = LineCnt + 2
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        For z = 1 To PropTransCnt
          Get TTHandle, ThisPropTrans(z), TaxTrans
          If TaxTrans.BelongTo = BillRec Then
            PrincDif = 0
            IntDif = 0
            AdvDif = 0
            LateListDif = 0
            Opt1Dif = 0
            Opt2Dif = 0
            Opt3Dif = 0
            GoSub GetType
            TotPropCnt = TotPropCnt + 1
            Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Principal    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(55); Using("$##,##0.00", TaxTrans.Revenue.Principle1Pd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Interest     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Advertising  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            Print #RptHandle, Tab(20); "Late Listing "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              Print #RptHandle, RealPers
              Print #RptHandle, String(79, "-")
              LineCnt = LineCnt + 2
              GoSub PrintBillHeader
              Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
              LineCnt = LineCnt + 1
            End If
            If Len(QPTrim$(Opt1Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            If Len(QPTrim$(Opt2Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            If Len(QPTrim$(Opt3Desc)) > 0 Then
              Print #RptHandle, Tab(20); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(55); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                Print #RptHandle, RealPers
                Print #RptHandle, String(79, "-")
                LineCnt = LineCnt + 2
                GoSub PrintBillHeader
                Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); ThisType; Tab(35); TaxTrans.TaxYear; Tab(55); Using$("$##,##0.00", TaxTrans.Amount)
                LineCnt = LineCnt + 1
              End If
            End If
            Print #RptHandle, String(79, "-")
            Print #RptHandle,
            LineCnt = LineCnt + 2
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
        Next z
      End If
    Next x
MoveOn:
  Next y
  Unload FrmShowPctComp
  EnableCloseButton Me.hwnd, True
  Print #RptHandle, FF$
  Close

  If TotPropCnt = 0 Then
    Call TaxMsg(800, "There are no transactions saved containing the pointers necessary to generate this report.")
    Exit Sub
  End If

  ViewPrint RptFile, "Tax Customer History By Property", True

  Exit Sub

GetType:
  Select Case TaxTrans.TranType
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit Applied at Billing"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Overpayment"
    Case 22
      ThisType = "Paid Bill Plus Overpay"
'    Case 23
'      ThisType = "Adjust Bill Down Affecting Credit Balance"
    Case 24
      ThisType = "Adjust Bill Up Affecting Credit Balance"
    Case Else
      ThisType = "Unknown"
  End Select

  Return

PrintBillHeader:
  Print #RptHandle, "Bill Number: " + ThisBillNum; Tab(30); "Property Pin #: " + ThisPin
  Print #RptHandle, String(79, ".")
  LineCnt = LineCnt + 2
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Tax Customer Transaction History By Property"
  Print #RptHandle, Tab(30); "Total Balance: " + Using$("$##,##0.00", TotBal)
  Print #RptHandle, Town$; Tab(71); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "For : #" + QPTrim$(Using$("####0", TaxCust.PIN)) + "  " + QPTrim$(TaxCust.CustName)
  Print #RptHandle, "Property Type"
  Print #RptHandle, "Trans Date"; Tab(12); "Trans Type"; Tab(35); "Tax Year"; Tab(56); "Trans Amt"; Tab(72); "Bill Bal"
  Print #RptHandle, Tab(20); "Revenue Type"; Tab(36); "Amount Billed"; Tab(54); "Amount Paid"; Tab(69); "Revenue Bal"
  Print #RptHandle, String(79, "-")
  LineCnt = 9

  Return
  
End Sub
