VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustByCat 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customers By Category Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmBLCustByCat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6030
      Left            =   1913
      TabIndex        =   1
      Top             =   1335
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10636
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLCustByCat.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   0
         Tag             =   $"frmBLCustByCat.frx":08E6
         Top             =   1830
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
         ColDesigner     =   "frmBLCustByCat.frx":09BF
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   4
         Tag             =   $"frmBLCustByCat.frx":0CBA
         Top             =   3765
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
         ColDesigner     =   "frmBLCustByCat.frx":0D73
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   5040
         TabIndex        =   3
         Tag             =   "You can elect to include all inactive accounts on this report. Select 'Yes' in the drop down list to include inactive accounts."
         Top             =   3150
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
         ColDesigner     =   "frmBLCustByCat.frx":106E
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3315
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   4770
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
         ButtonDesigner  =   "frmBLCustByCat.frx":1369
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5190
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLCustByCat.frx":1547
         Top             =   4770
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
         ButtonDesigner  =   "frmBLCustByCat.frx":15F2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   405
         Left            =   4605
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLCustByCat.frx":17D1
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
         ButtonDesigner  =   "frmBLCustByCat.frx":188A
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   390
         Left            =   2730
         TabIndex        =   2
         Tag             =   $"frmBLCustByCat.frx":1A6E
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
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLCustByCat.frx":1B85
         Top             =   4770
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
         ButtonDesigner  =   "frmBLCustByCat.frx":1C55
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   1005
         Top             =   1485
         Width           =   5970
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Customers By Category Report"
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
         TabIndex        =   13
         Top             =   576
         Width           =   4332
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
         TabIndex        =   12
         Top             =   2595
         Width           =   1350
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
         Left            =   1155
         TabIndex        =   11
         Top             =   3870
         Width           =   1500
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
         Height          =   345
         Left            =   1770
         TabIndex        =   10
         Top             =   3240
         Width           =   3030
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
         Height          =   255
         Left            =   1050
         TabIndex        =   9
         Top             =   5445
         Width           =   2100
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   2010
      TabIndex        =   15
      Top             =   7557
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
      Height          =   6300
      Left            =   1793
      Top             =   1215
      Width           =   8055
   End
End
Attribute VB_Name = "frmBLCustByCat"
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
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustByCat.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim x As Double, cnt As Double
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
  Dim TCat$, CustNum$
  Dim RptHandle As Integer
  Dim ThisCustCnt As Double
  Dim CatCode$, ThisCat$
  Dim dlm$, Nextx As Double
  Dim TownName$, NextCust As Double
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim InActiveFlag As Boolean
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim NameFlag As Boolean, NumFlag As Boolean
  Dim CatCnt As Integer, NumOfActCust As Integer
  Dim NumOfInActCust As Integer
  
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
  
  ReportFile$ = "BLRPTS\ARCustByCat.RPT"
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
  
  For x = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(x), CustRec
    If CustRec.Inactive = "N" Then
      NumOfActCust = NumOfActCust + 1
    End If
  Next x
  
  OpenCatCodeIdxFile CodeIdxHandle
  NumOfARCatRecs = LOF(CodeIdxHandle) / Len(CodeIdxRec)
  ReDim CodeIdx(1 To NumOfARCatRecs) As Integer
  For x = 1 To NumOfARCatRecs
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  
  OpenCatCodeFile CHandle
  
  OpenCustFile CustHandle
  
  ReDim Category(1 To NumOfARCatRecs) As Integer
  ReDim CustByCat(1 To 1) As Integer
  ReDim CustByCatCnt(1 To NumOfARCatRecs) As Integer
  CustCnt = 0
  ThisCustCnt = 0
  
  frmBLShowPctComp.Label1 = "Loading Customer By Category Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For x = 1 To NumOfARCatRecs
    Get CHandle, CodeIdx(x), CodeRec
    If fptxtCatCode.Text <> "ALL" Then
      If QPTrim$(CodeRec.CatCode) <> fptxtCatCode.Text Then GoTo NotALL
    End If
    ThisCat = QPTrim$(CodeRec.CatCode)
    For cnt = 1 To NumOfCustIdxRecs
      Get CustHandle, IdxRecs(cnt), CustRec
      If InActiveFlag = False Then
        If QPTrim$(CustRec.Inactive) = "Y" Then
          GoTo NoneHere
        End If
      End If
      If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
        If QPTrim$(CustRec.BILLCAT1) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT2) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT3) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT4) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT5) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        End If
      End If
NoneHere:
    Next cnt
    CustByCatCnt(x) = ThisCustCnt
    If ThisCustCnt > 0 Then CatCnt = CatCnt + 1
    ThisCustCnt = 0
    frmBLShowPctComp.ShowPctComp x, NumOfARCatRecs
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
NotALL:
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If CustCnt = 0 Then
    Close         'Close all open files now
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers on file that fit the criteria entered."
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  CustCnt = 1
  For x = 1 To NumOfARCatRecs
    Get CHandle, CodeIdx(x), CodeRec
      For cnt = 1 To CustByCatCnt(x)
        Print #RptHandle, TownName; dlm; TCat$; dlm;
        Print #RptHandle, QPTrim$(CodeRec.CatCode); dlm;
        Print #RptHandle, QPTrim$(CodeRec.CODEDESC); dlm;
        Get CustHandle, CustByCat(CustCnt), CustRec
        Print #RptHandle, QPTrim$(CustRec.CustNumb); dlm; QPTrim$(CustRec.BillName); dlm; QPTrim$(CustRec.CustName); dlm;
        If fptxtCatCode.Text = "ALL" Then
          If InActiveFlag = False Then
            Print #RptHandle, CStr(NumOfActCust); dlm;
          Else
            Print #RptHandle, CStr(NumOfCustIdxRecs); dlm;
          End If
        Else
          Print #RptHandle, ""; dlm;
        End If
        If fptxtCatCode.Text = "ALL" Then
          Print #RptHandle, CStr(CatCnt); dlm;
        Else
          Print #RptHandle, ""; dlm;
        End If
        If CustRec.Inactive = "Y" Then
          Print #RptHandle, "InActive"
        Else
          Print #RptHandle, ""
        End If
        CustCnt = CustCnt + 1
      Next cnt
  Next x
  
  Close         'Close all open files now
  
  
  arBLCustByCat.Show
  frmBLLoadReport.Show
  
  MainLog ("The 'Customer By Category Report' was processed in graphical format.")
  
  Exit Sub
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustByCat", "PrintGraphics", Erl)
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
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim TCat$, CustNum$
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim CatCode$
  Dim InActiveFlag As Boolean
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisCustCnt As Double
  Dim TownName$
  Dim ThisCat$, NumOfActCust As Integer
  Dim Nextx As Double
  Dim NextCust As Double
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim Active$, cnt As Integer
  Dim ThisCust As String * 26
  Dim ThisBill As String * 35
  Dim CatCount As Integer
  
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

  ReportFile$ = "ARCustByCat.PRN"
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

  For x = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(x), CustRec
    If CustRec.Inactive = "N" Then
      NumOfActCust = NumOfActCust + 1
    End If
  Next x

  OpenCatCodeIdxFile CodeIdxHandle
  NumOfARCatRecs = LOF(CodeIdxHandle) / Len(CodeIdxRec)
  ReDim CodeIdx(1 To NumOfARCatRecs) As Integer
  For x = 1 To NumOfARCatRecs
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle
  
  ReDim Category(1 To NumOfARCatRecs) As Integer
  ReDim CustByCat(1 To 1) As Integer
  ReDim CustByCatCnt(1 To NumOfARCatRecs) As Integer
  
  CustCnt = 0
  ThisCustCnt = 0
  GoSub PrintHeader
  
  frmBLShowPctComp.Label1 = "Loading Customers By Category Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  For x = 1 To NumOfARCatRecs
    Get CHandle, CodeIdx(x), CodeRec
    If fptxtCatCode.Text <> "ALL" Then
      If QPTrim$(CodeRec.CatCode) <> fptxtCatCode.Text Then GoTo NoneHere
    End If
    Category(x) = CodeIdx(x)
    ThisCat = QPTrim$(CodeRec.CatCode)
    For cnt = 1 To NumOfCustIdxRecs
      Get CustHandle, IdxRecs(cnt), CustRec
      If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
        If InActiveFlag = False Then
          If QPTrim$(CustRec.Inactive) = "Y" Then
            GoTo NoSir
          End If
        End If
        If QPTrim$(CustRec.BILLCAT1) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT2) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT3) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT4) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        ElseIf QPTrim$(CustRec.BILLCAT5) = ThisCat Then
          CustCnt = CustCnt + 1
          ThisCustCnt = ThisCustCnt + 1
          ReDim Preserve CustByCat(1 To CustCnt) As Integer
          CustByCat(CustCnt) = IdxRecs(cnt)
        End If
      End If
NoSir:
    Next cnt
    
    CustByCatCnt(x) = ThisCustCnt
    ThisCustCnt = 0
    frmBLShowPctComp.ShowPctComp x, NumOfARCatRecs
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
    If fptxtCatCode.Text <> "ALL" Then Exit For
NoneHere:
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If CustCnt = 0 Then
    Close         'Close all open files now
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Label1.Caption = "There are no customers on file that fit the criteria entered."
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  CustCnt = 1
  For x = 1 To NumOfARCatRecs
    Get CHandle, CodeIdx(x), CodeRec
      If CustByCatCnt(x) = 0 Then GoTo NoneHere2
      CatCount = CatCount + 1
      Print #RptHandle, QPTrim$(CodeRec.CatCode); "  ";
      Print #RptHandle, QPTrim$(CodeRec.CODEDESC)
      Print #RptHandle, "Customer Number"; Tab(24); "Billing Name"; Tab(60); "Customer Name"
      Print #RptHandle, String$(85, "-")
      LineCnt = LineCnt + 3
      For cnt = 1 To CustByCatCnt(x)
        Get CustHandle, CustByCat(CustCnt), CustRec
        ThisBill = QPTrim$(CustRec.BillName)
        ThisCust = QPTrim$(CustRec.CustName)
        If CustRec.Inactive = "Y" Then
          Active$ = " *INACTIVE* "
        Else
          Active$ = ""
        End If
        If IsNumeric(CustRec.CustNumb) Then
          Print #RptHandle, Using$("##########", CDbl(CustRec.CustNumb)); Tab(11); Active$; Tab(24); ThisBill; Tab(60); ThisCust
        Else
          Print #RptHandle, QPTrim$(CustRec.CustNumb); Tab(11); Active$; Tab(24); ThisBill; Tab(60); ThisCust
        End If
        LineCnt = LineCnt + 1
        If LineCnt > (MaxLines - 3) Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          Print #RptHandle, QPTrim$(CodeRec.CatCode); "  ";
          Print #RptHandle, QPTrim$(CodeRec.CODEDESC)
          Print #RptHandle, "Customer Number"; Tab(24); "Billing Name"; Tab(60); "Customer Name"
          Print #RptHandle, String$(85, "-")
          LineCnt = LineCnt + 3
        End If
        CustCnt = CustCnt + 1
      Next cnt
      If LineCnt <> 5 And LineCnt <> 7 Then
        Print #RptHandle, String$(85, "-")
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Tab(8); QPTrim$(CodeRec.CODEDESC); Tab(45); "# of Customers: " + Using$("#####", CustByCatCnt(x))
      LineCnt = LineCnt + 1
      Print #RptHandle, String$(85, "=")
      Print #RptHandle,
      Print #RptHandle,
      LineCnt = LineCnt + 3
      If LineCnt > (MaxLines - 3) Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        If CustByCatCnt(x) <> cnt - 1 Then
          Print #RptHandle, QPTrim$(CodeRec.CatCode); "  ";
          Print #RptHandle, QPTrim$(CodeRec.CODEDESC)
          Print #RptHandle, "Customer Number"; Tab(21); "Billing Name"; Tab(57); "Customer Name"
          Print #RptHandle, String$(85, "-")
          LineCnt = LineCnt + 3
        End If
      End If
NoneHere2:
  Next x
  
  If fptxtCatCode.Text = "ALL" Then
    GoSub PrintEnding
  End If
  Print #RptHandle, FF$ ' Chr$(18);   ' oki 320 10 cpi
  
  Close
  
  ViewPrint ReportFile$, "Customers By Category Report", True
  
  KillFile ReportFile$
  
  MainLog ("The 'Customer By Category Report' was processed in text format.")
  
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License: Customers By Category Report"
  Print #RptHandle, TownName
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Str(Page)
  If TCat$ = "ALL" Then
    Print #RptHandle, "Category: " + TCat$
  Else
    Print #RptHandle, "Category: " + TCat$ + "/" + GetCatDesc(TCat$)
  End If
  Print #RptHandle, String$(85, "=")
  LineCnt = 5

  Return

PrintEnding:
  If InActiveFlag = False Then
    Print #RptHandle, "Total Categories with Customers: " + Using("####0", CatCount) + "  " + "Total Customers Printed: "; Using("####0", NumOfActCust)
  Else
    Print #RptHandle, "Total Categories with Customers: " + Using("####0", CatCount) + "  " + "Total Customers Printed: "; Using("####0", NumOfCustIdxRecs)
  End If
  
  Return

End Sub
Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  
  lblBalloon.Visible = False
  One = 1
  DHandle = FreeFile
  Open "custByCat.dat" For Output As DHandle Len = 2
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
  KillFile "custByCat.dat"
  DoEvents
  Unload frmBLCustByCat
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
    frmBLMessageBoxJr.Label1.Caption = "Pitch 12 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
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
    fpcmbIncInactive.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub fptxtCatCode_Change()
  If QPTrim$(fptxtCatCode.Text) = "" Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub

Private Sub fptxtCatCode_LostFocus()
  If Not IsNumeric(fptxtCatCode.Text) Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub
