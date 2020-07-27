VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLInOutRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Inside/Outside City Limits Report"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLInOutRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleMode       =   0  'User
   ScaleWidth      =   11724
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5892
      Left            =   1969
      TabIndex        =   4
      Top             =   1425
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   10393
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLInOutRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   3075
         TabIndex        =   3
         Tag             =   $"frmBLInOutRpt.frx":08E6
         Top             =   3795
         Width           =   3555
         _Version        =   196608
         _ExtentX        =   6271
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
         ColDesigner     =   "frmBLInOutRpt.frx":099F
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   0
         Tag             =   $"frmBLInOutRpt.frx":0C9A
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
         ColDesigner     =   "frmBLInOutRpt.frx":0D32
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   5085
         TabIndex        =   2
         Tag             =   "You can elect to include inactive accounts in this report by selecting 'Yes' from the drop down box."
         Top             =   3165
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
         ColDesigner     =   "frmBLInOutRpt.frx":102D
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   3168
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Business License Reports' menu."
         Top             =   4752
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
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
         ButtonDesigner  =   "frmBLInOutRpt.frx":1328
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   630
         Left            =   5190
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLInOutRpt.frx":1506
         Top             =   4755
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1111
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
         ButtonDesigner  =   "frmBLInOutRpt.frx":1682
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   405
         Left            =   4605
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLInOutRpt.frx":1861
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
         ButtonDesigner  =   "frmBLInOutRpt.frx":191A
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   390
         Left            =   2685
         TabIndex        =   1
         Tag             =   $"frmBLInOutRpt.frx":1AFE
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
         Height          =   630
         Left            =   870
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLInOutRpt.frx":1C14
         Top             =   4755
         Width           =   2160
         _Version        =   131072
         _ExtentX        =   3810
         _ExtentY        =   1111
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
         ButtonDesigner  =   "frmBLInOutRpt.frx":1CE4
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
         Left            =   1296
         TabIndex        =   14
         Top             =   3888
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
         TabIndex        =   13
         Top             =   2595
         Width           =   1350
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1200
         Top             =   432
         Width           =   5580
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Inside/Outside City Limits Report"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   576
         Width           =   5148
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
         Top             =   1920
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3036
         Left            =   1008
         Top             =   1488
         Width           =   5964
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
         Left            =   1824
         TabIndex        =   10
         Top             =   3264
         Width           =   3036
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
         Left            =   912
         TabIndex        =   9
         Top             =   5424
         Width           =   2100
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   720
      TabIndex        =   15
      Top             =   2538
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   794
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
      Height          =   6156
      Left            =   1801
      Top             =   1287
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLInOutRpt"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLInOutRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Public Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  
  lblBalloon.Visible = False
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  fptxtCatCode.ToolTipText = "You can select ALL or you can select a specific category for which to print this report."
'  cmdCodeList.ToolTipText = "Press to bring up a complete category list."
'  fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include all inactive accounts."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  cmdExit.ToolTipText = ""
'  cmdProcess.ToolTipText = "Press to activate this report."
  One = 1
  DHandle = FreeFile
  Open "inoutrpt.dat" For Output As DHandle Len = 2
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

Private Sub fpcmbIncInactive_Change()
  If QPTrim$(fpcmbIncInactive.Text) = "" Then
    fpcmbIncInactive.Text = "No"
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
  KillFile "inoutrpt.dat"
  DoEvents
  Unload frmBLInOutRpt
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
    frmBLMessageBoxJr.Label1.Caption = "Pitch 17 is recommended for this report."
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
'    fpcmbIncInactive.ToolTipText = "Choose 'Yes' if you wish to include all inactive accounts."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    cmdExit.ToolTipText = ""
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
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim LicNumIdxRec As CustLicNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim LicFlag As Boolean
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim ZCnt&, cnt&
  Dim InActiveFlag As Boolean
  Dim WhereFlag As Integer
  Dim LoopCnt As Integer
  Dim CustFee#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim CatCode$
  Dim Snt&
  Dim CodeHandle As Integer
  Dim TotCustFees#
  Dim TotOutBal#
  Dim AveBal#
  Dim AveFee#
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim GrandBal#
  Dim GrandFees#
  Dim GBalAve#
  Dim GFeeAve#
  Dim GCustCnt As Integer
  Dim TCat$
  Dim PctCnt As Long
  Dim ThisCode$
  
  On Error GoTo ERRORSTUFF
  
  TCat$ = QPTrim$(fptxtCatCode.Text)
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  
  If QPTrim$(fptxtCatCode.Text) = "ALL" Then
    ThisCode = "ALL"
  Else
    For x = 1 To NumOfARCatRecs
      Get CodeHandle, x, CodeRec
      If QPTrim$(CodeRec.CatCode) = QPTrim$(fptxtCatCode.Text) Then
        ThisCode$ = QPTrim$(CodeRec.CODEDESC)
        Exit For
      End If
    Next x
  End If
  
  WhereFlag = 1
  LoopCnt = 0
  InActiveFlag = False
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  
  ReportFile$ = "ARIORPT.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  NameFlag = False
  NumFlag = False
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
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
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
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
  ElseIf NumFlag = True Then
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
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  ElseIf NumFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  
  Close IdxHandle
  GoSub PrintInOutRptHeader
  frmBLShowPctComp.Label1 = "Loading In/Out City Limits Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  TotOutBal# = 0
  TotCustFees# = 0
  
  ReDim GTotBal(1 To 3) As Double
  ReDim GTotFees(1 To 3) As Double
  ReDim CustCnt(1 To 3) As Integer
  Do
    For ZCnt& = 1 To NumOfCustIdxRecs
      Get CustHandle, IdxRecs(ZCnt), CustRec
      'user can include inactive accounts
      If InActiveFlag = False Then
        If QPTrim$(CustRec.Inactive) = "Y" Then
          GoTo Inactive
        End If
      End If
      If TCat$ <> RTrim$(CustRec.BILLCAT1) And TCat$ <> RTrim$(CustRec.BILLCAT2) And TCat$ <> RTrim$(CustRec.BILLCAT3) And TCat$ <> RTrim$(CustRec.BILLCAT4) And TCat$ <> RTrim$(CustRec.BILLCAT5) And TCat$ <> "ALL" Then GoTo Inactive
      If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
        If LoopCnt = 0 Then
          If CustRec.CustLocation <> "I" Then GoTo Inactive
        ElseIf LoopCnt = 1 Then
          If CustRec.CustLocation <> "O" Then GoTo Inactive
        ElseIf LoopCnt = 2 Then
          If CustRec.CustLocation <> " " Then GoTo Inactive
        End If
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintInOutRptHeader
        End If
        GoSub GetCustFee
        GTotBal(WhereFlag) = GTotBal(WhereFlag) + CustRec.AcctBal
        GTotFees(WhereFlag) = GTotFees(WhereFlag) + CustFee#
        If QPTrim$(CustRec.LICENSE) = "" Then CustRec.LICENSE = "0"
        Print #RptHandle, IdxRecs(ZCnt); Tab(10); CustRec.BillName; Tab(46); Using$("###########0", CustRec.LICENSE); Tab(62); MakeRegDate(CustRec.VALID); Tab(75); Using$("$#,###,##0.00", CustRec.AcctBal); Tab(95); Using$("$#,###,##0.00", CustFee#)
        CustCnt(WhereFlag) = CustCnt(WhereFlag) + 1
        LineCnt = LineCnt + 1
      End If

Inactive:
      PctCnt = PctCnt + 1
      frmBLShowPctComp.ShowPctComp PctCnt, NumOfCustIdxRecs * 3
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
    Next ZCnt&
    GoSub PrintInOutRptEnding
    WhereFlag = WhereFlag + 1
    If WhereFlag = 4 Then Exit Do
    GoSub PrintInOutRptHeader
    LoopCnt = LoopCnt + 1
  Loop
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True

  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  If CustCnt(1) + CustCnt(2) + CustCnt(3) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers in category " + TCat$ + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  ViewPrint ReportFile$, "Inside/Outside City Limits Report", True
  
  KillFile ReportFile$
  
  MainLog ("'Inside/Outside City Limits Report' processed in text format.")
  
  Exit Sub
  
PrintInOutRptHeader:
  Page = Page + 1
  If WhereFlag = 1 Then  'inside
    Print #RptHandle, Tab(20); "Business Licenses: Inside City Limits"
  ElseIf WhereFlag = 2 Then 'outside
    Print #RptHandle, Tab(20); "Business Licenses: Outside City Limits"
  ElseIf WhereFlag = 3 Then 'not known
    Print #RptHandle, Tab(20); "Business Licenses: Location Not Saved"
  End If
  Print #RptHandle, QPTrim$(TownRec.TownName); Tab(65); "Page #"; Page
  If TownRec.IssFee > 0 Then
    Print #RptHandle, "Category Code: " + QPTrim$(fptxtCatCode.Text) + "/" + ThisCode; Tab(65); "Current fees include a " + QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) + " issuance fee."
  Else
    Print #RptHandle, "Category Code: " + QPTrim$(fptxtCatCode.Text) + "/" + ThisCode
  End If
  Print #RptHandle, "Report Date: "; Date$
  Print #RptHandle, "Cust #"; Tab(10); "Billing Name"; Tab(50); "License #"; Tab(62); "Valid Thru"; Tab(75); "Total Balance"; Tab(91); "Current Total Fee"
  Print #RptHandle, String$(107, "=")
  LineCnt = 6
  Return
  
PrintInOutRptEnding:
  If CustCnt(WhereFlag) > 0 Then
    AveBal# = OldRound(GTotBal(WhereFlag) / CustCnt(WhereFlag))
  Else
    AveBal# = 0
  End If
  If CustCnt(WhereFlag) > 0 Then
    AveFee# = OldRound(GTotFees(WhereFlag) / CustCnt(WhereFlag))
  Else
    AveFee# = 0
  End If
    
  If LineCnt >= MaxLines - 12 Then
    GoSub PrintInOutRptHeader
    If WhereFlag = 1 Then
      Print #RptHandle, "Business Licenses: Inside City Limits"
    ElseIf WhereFlag = 2 Then
      Print #RptHandle, "Business Licenses: Outside City Limits"
    ElseIf WhereFlag = 3 Then
      Print #RptHandle, "Business Licenses: Location Not Saved"
    End If
    Print #RptHandle, "Customers Counted: "; Using("#####0", CustCnt(WhereFlag)); Tab(75); Using$("$#,###,##0.00", GTotBal(WhereFlag)); Tab(95); Using$("$#,###,##0.00", GTotFees(WhereFlag))
    Print #RptHandle, "Average Balance = " + QPTrim$(Using("$#,###,##0.00", AveBal#)); Tab(40); "Average Fee = " + QPTrim$(Using("$#,###,##0.00", AveFee#))
    LineCnt = LineCnt + 4
    GoSub PrintGrandTotal
    Print #RptHandle, FF$
  Else
    Print #RptHandle, String$(107, "-")
    If WhereFlag = 1 Then
      Print #RptHandle, "Business Licenses: Inside City Limits"
    ElseIf WhereFlag = 2 Then
      Print #RptHandle, "Business Licenses: Outside City Limits"
    ElseIf WhereFlag = 3 Then
      Print #RptHandle, "Business Licenses: Location Not Saved"
    End If
    Print #RptHandle, "Customers Counted: "; Using("#####0", CustCnt(WhereFlag)); Tab(75); Using$("$#,###,##0.00", GTotBal(WhereFlag)); Tab(95); Using$("$#,###,##0.00", GTotFees(WhereFlag))
    Print #RptHandle, "Average Balance = "; Tab(20); Using("$#,###,##0.00", AveBal#); Tab(40); "Average Fee = "; Tab(56); Using("$#,###,##0.00", AveFee#)
    LineCnt = LineCnt + 5
    If WhereFlag = 3 Then
      If LineCnt >= MaxLines - 12 Then
        Print #RptHandle, FF$
        GoSub PrintGrandTotal
      Else
        GoSub PrintGrandTotal
      End If
    End If
    Print #RptHandle, FF$
  End If
  Return

PrintGrandTotal:
  GrandFees# = OldRound(GTotFees(1) + GTotFees(2) + GTotFees(3))
  GrandBal# = OldRound(GTotBal(1) + GTotBal(2) + GTotBal(3))
  GCustCnt = CustCnt(1) + CustCnt(2) + CustCnt(3)
  GBalAve = GrandBal#
  If GCustCnt > 0 Then
    GBalAve = OldRound(GBalAve / GCustCnt)
  Else
    GBalAve = 0
  End If
  
  GFeeAve# = GrandFees#
  If GCustCnt > 0 Then
    GFeeAve# = OldRound(GFeeAve / GCustCnt)
  Else
    GFeeAve# = 0
  End If
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, String$(107, "=")
  Print #RptHandle, "Summary Of Inside/Outside Report"
  Print #RptHandle,
  Print #RptHandle, Tab(10); "Total Customers"; Tab(27); "Total Balance"; Tab(45); "Balance Ave"; Tab(62); "Total Fees"; Tab(79); "Fees Ave"
  Print #RptHandle, String$(86, "-")
  If CustCnt(1) > 0 Then
    Print #RptHandle, "Inside"; Tab(15); Using("####", CustCnt(1)); Tab(26); Using("$##,###,##0.00", GTotBal(1)); Tab(43); Using("$#,###,##0.00", OldRound(GTotBal(1) / CustCnt(1))); Tab(58); Using("$##,###,##0.00", GTotFees(1)); Tab(74); Using("$#,###,##0.00", OldRound(GTotFees(1) / CustCnt(1)))
  Else
    Print #RptHandle, "Inside"; Tab(15); Using("####", 0); Tab(26); Using("$##,###,##0.00", 0); Tab(43); Using("$#,###,##0.00", 0); Tab(58); Using("$##,###,##0.00", 0); Tab(74); Using("$#,###,##0.00", 0)
  End If
  
  If CustCnt(2) > 0 Then
    Print #RptHandle, "Outside"; Tab(15); Using("####", CustCnt(2)); Tab(26); Using("$##,###,##0.00", GTotBal(2)); Tab(43); Using("$#,###,##0.00", OldRound(GTotBal(2) / CustCnt(2))); Tab(58); Using("$##,###,##0.00", GTotFees(2)); Tab(74); Using("$#,###,##0.00", OldRound(GTotFees(2) / CustCnt(2)))
  Else
    Print #RptHandle, "Outside"; Tab(15); Using("####", 0); Tab(26); Using("$##,###,##0.00", 0); Tab(43); Using("$#,###,##0.00", 0); Tab(58); Using("$##,###,##0.00", 0); Tab(74); Using("$#,###,##0.00", 0)
  End If
  
  If CustCnt(3) > 0 Then
    Print #RptHandle, "Not Saved"; Tab(15); Using("####", CustCnt(3)); Tab(26); Using("$##,###,##0.00", GTotBal(3)); Tab(43); Using("$#,###,##0.00", OldRound(GTotBal(3) / CustCnt(3))); Tab(58); Using("$##,###,##0.00", GTotFees(3)); Tab(74); Using("$#,###,##0.00", OldRound(GTotFees(3) / CustCnt(3)))
  Else
    Print #RptHandle, "Not Saved"; Tab(15); Using("####", 0); Tab(26); Using("$##,###,##0.00", 0); Tab(43); Using("$#,###,##0.00", 0); Tab(58); Using("$##,###,##0.00", 0); Tab(74); Using("$#,###,##0.00", 0)
  End If
  
  Print #RptHandle, String$(86, "-")
  Print #RptHandle, "Totals"; Tab(14); Using("#####", GCustCnt); Tab(26); Using("$##,###,##0.00", GrandBal#); Tab(43); Using("$#,###,##0.00", GBalAve);
  Print #RptHandle, Tab(58); Using("$##,###,##0.00", GrandFees#); Tab(74); Using("$#,###,##0.00", GFeeAve#)
  Return

GetCustFee:
  
  CustFee# = 0
  FeeAmt1# = 0
  FeeAmt2# = 0
  FeeAmt3# = 0
  FeeAmt4# = 0
  FeeAmt5# = 0
  
  Prorate# = CustRec.Prorate
  
  If Prorate# >= 100 Or Prorate# < 0 Then
    Prorate# = 1
  Else
    Prorate# = OldRound#(Prorate# * 0.01)
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        CustRec.DESC1 = CodeRec.CODEDESC           'Reset Code Descriptions
        If CodeRec.CodeType = "F" Then
          FeeAmt1# = CodeRec.Fee
          FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV1
          FeeAmt1# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV1
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt1# < CodeRec.BaseAmt1 Then FeeAmt1# = CodeRec.BaseAmt1
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt1# < CodeRec.BaseAmt2 Then FeeAmt1# = CodeRec.BaseAmt2
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt1# < CodeRec.BaseAmt3 Then FeeAmt1# = CodeRec.BaseAmt3
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt1# < CodeRec.BaseAmt4 Then FeeAmt1# = CodeRec.BaseAmt4
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt1# < CodeRec.BaseAmt5 Then FeeAmt1# = CodeRec.BaseAmt5
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt1# < CodeRec.BaseAmt6 Then FeeAmt1# = CodeRec.BaseAmt6
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C2:             'Catagory #2
  
'  CustFee# = OldRound#(CustFee# + FeeAmt1#)
'  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt2# = CodeRec.Fee
          FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV2
          FeeAmt2# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV2
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt2# < CodeRec.BaseAmt1 Then FeeAmt2# = CodeRec.BaseAmt1
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt2# < CodeRec.BaseAmt2 Then FeeAmt2# = CodeRec.BaseAmt2
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt2# < CodeRec.BaseAmt3 Then FeeAmt2# = CodeRec.BaseAmt3
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt2# < CodeRec.BaseAmt4 Then FeeAmt2# = CodeRec.BaseAmt4
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt2# < CodeRec.BaseAmt5 Then FeeAmt2# = CodeRec.BaseAmt5
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt2# < CodeRec.BaseAmt6 Then FeeAmt2# = CodeRec.BaseAmt6
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C3:
'  CustFee# = OldRound#(CustFee# + FeeAmt2#)
'  FeeAmt# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt3# = CodeRec.Fee
          FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV3
          FeeAmt3# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV3
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt3# < CodeRec.BaseAmt1 Then FeeAmt3# = CodeRec.BaseAmt1
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt3# < CodeRec.BaseAmt2 Then FeeAmt3# = CodeRec.BaseAmt2
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt3# < CodeRec.BaseAmt3 Then FeeAmt3# = CodeRec.BaseAmt3
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt3# < CodeRec.BaseAmt4 Then FeeAmt3# = CodeRec.BaseAmt4
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt3# < CodeRec.BaseAmt5 Then FeeAmt3# = CodeRec.BaseAmt5
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt3# < CodeRec.BaseAmt6 Then FeeAmt3# = CodeRec.BaseAmt6
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
           GoTo c4
          End If
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 3
  
c4:
'  CustFee# = OldRound#(CustFee# + FeeAmt3#)
'  FeeAmt4# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt4# = CodeRec.Fee
          FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV4
          FeeAmt4# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV4
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt4# < CodeRec.BaseAmt1 Then FeeAmt4# = CodeRec.BaseAmt1
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt4# < CodeRec.BaseAmt2 Then FeeAmt4# = CodeRec.BaseAmt2
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt4# < CodeRec.BaseAmt3 Then FeeAmt4# = CodeRec.BaseAmt3
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt4# < CodeRec.BaseAmt4 Then FeeAmt4# = CodeRec.BaseAmt4
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt4# < CodeRec.BaseAmt5 Then FeeAmt4# = CodeRec.BaseAmt5
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt4# < CodeRec.BaseAmt6 Then FeeAmt4# = CodeRec.BaseAmt6
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
c5:
'  CustFee# = OldRound#(CustFee# + FeeAmt4#)
'  FeeAmt5# = 0
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt5# = CodeRec.Fee
          FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV5
          FeeAmt5# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV5
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt5# < CodeRec.BaseAmt1 Then FeeAmt5# = CodeRec.BaseAmt1
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt5# < CodeRec.BaseAmt2 Then FeeAmt5# = CodeRec.BaseAmt2
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt5# < CodeRec.BaseAmt3 Then FeeAmt5# = CodeRec.BaseAmt3
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt5# < CodeRec.BaseAmt4 Then FeeAmt5# = CodeRec.BaseAmt4
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt5# < CodeRec.BaseAmt5 Then FeeAmt5# = CodeRec.BaseAmt5
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt5# < CodeRec.BaseAmt6 Then FeeAmt5# = CodeRec.BaseAmt6
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          
        End If
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
SkipEm:
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + TownRec.IssFee)
'  FeeAmt# = 0
  
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicListRpt", "PrintText", Erl)
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
  Dim ReportFile$
  Dim x As Double
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim LicNumIdxRec As CustLicNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim ZCnt&, cnt&
  Dim InActiveFlag As Boolean
  Dim WhereFlag As Integer
  Dim LoopCnt As Integer
  Dim CustFee#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim CatCode$
  Dim Snt&
  Dim CodeHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim GrandBal#
  Dim GrandFees#
  Dim GBalAve#
  Dim GFeeAve#
  Dim GCustCnt As Integer
  Dim dlm$, TCat$
  Dim PctCnt As Long
  Dim ThisCode$
  
  On Error GoTo ERRORSTUFF
  
  TCat$ = QPTrim$(fptxtCatCode.Text)
  dlm$ = "~"
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  If QPTrim$(fptxtCatCode.Text) = "ALL" Then
    ThisCode = "ALL"
  Else
    For x = 1 To NumOfARCatRecs
      Get CodeHandle, x, CodeRec
      If QPTrim$(CodeRec.CatCode) = QPTrim$(fptxtCatCode.Text) Then
        ThisCode$ = QPTrim$(CodeRec.CODEDESC)
        Exit For
      End If
    Next x
  End If
  
  WhereFlag = 1
  LoopCnt = 0
  InActiveFlag = False
  If QPTrim$(fpcmbIncInactive.Text) = "Yes" Then
    InActiveFlag = True
  End If
  
  ReportFile$ = "BLRPTS\ARIORPT.RPT"  'Report File Name
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  NameFlag = False
  NumFlag = False
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
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
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
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
  ElseIf NumFlag = True Then
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
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  ElseIf NumFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
  frmBLShowPctComp.Label1 = "Loading In/Out City Limits Report"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  ReDim GTotBal(1 To 3) As Double
  ReDim GTotFees(1 To 3) As Double
  ReDim CustCnt(1 To 3) As Integer
  ReDim AveBal(1 To 3) As Double
  ReDim AveFee(1 To 3) As Double
  PctCnt = 1
  Do
    For ZCnt& = 1 To NumOfCustIdxRecs
      Get CustHandle, IdxRecs(ZCnt), CustRec
      'user can include inactive accounts
      If InActiveFlag = False Then
        If QPTrim$(CustRec.Inactive) = "Y" Then
          GoTo Inactive
        End If
      End If
      If TCat$ <> RTrim$(CustRec.BILLCAT1) And TCat$ <> RTrim$(CustRec.BILLCAT2) And TCat$ <> RTrim$(CustRec.BILLCAT3) And TCat$ <> RTrim$(CustRec.BILLCAT4) And TCat$ <> RTrim$(CustRec.BILLCAT5) And TCat$ <> "ALL" Then GoTo Inactive
      If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" Then
        If LoopCnt = 0 Then
          If CustRec.CustLocation <> "I" Then GoTo Inactive
        ElseIf LoopCnt = 1 Then
          If CustRec.CustLocation <> "O" Then GoTo Inactive
        ElseIf LoopCnt = 2 Then
          If CustRec.CustLocation <> " " Then GoTo Inactive
        End If
        GoSub GetCustFee
        GTotBal(WhereFlag) = GTotBal(WhereFlag) + CustRec.AcctBal
        CustCnt(WhereFlag) = CustCnt(WhereFlag) + 1
        GTotFees(WhereFlag) = GTotFees(WhereFlag) + CustFee#
        '                     0                1                2
        Print #RptHandle, CustCnt(1); dlm; CustCnt(2); dlm; CustCnt(3); dlm;
        '                     3                4                5
        Print #RptHandle, GTotBal(1); dlm; GTotBal(2); dlm; GTotBal(3); dlm;
        '                     6                  7                8
        Print #RptHandle, GTotFees(1); dlm; GTotFees(2); dlm; GTotFees(3); dlm;
        '                      9                     10                     11
        Print #RptHandle, IdxRecs(ZCnt); dlm; CustRec.BillName; dlm; CustRec.LICENSE; dlm;
        '                             12                         13                 14
        Print #RptHandle, MakeRegDate(CustRec.VALID); dlm; CustRec.AcctBal; dlm; CustFee#; dlm;
        '                            15                       16
        Print #RptHandle, QPTrim$(TownRec.TownName); dlm; WhereFlag; dlm;
        If WhereFlag = 1 Then
          '                                 17
          Print #RptHandle, "Business License: Inside City Limits"; dlm;
        ElseIf WhereFlag = 2 Then
          '                                 17
          Print #RptHandle, "Business License: Outside City Limits"; dlm;
        Else
          '                                 17
          Print #RptHandle, "Business License: Location Not Saved"; dlm;
        End If
      End If
      
      If CustCnt(WhereFlag) > 0 Then
        AveBal(WhereFlag) = OldRound(GTotBal(WhereFlag) / CustCnt(WhereFlag))
      Else
        AveBal(WhereFlag) = 0
      End If
      If CustCnt(WhereFlag) > 0 Then
        AveFee(WhereFlag) = OldRound(GTotFees(WhereFlag) / CustCnt(WhereFlag))
      Else
        AveFee(WhereFlag) = 0
      End If
      GrandFees# = OldRound(GTotFees(1) + GTotFees(2) + GTotFees(3))
      GrandBal# = OldRound(GTotBal(1) + GTotBal(2) + GTotBal(3))
      GCustCnt = CustCnt(1) + CustCnt(2) + CustCnt(3)
      GBalAve = GrandBal#
      If GCustCnt > 0 Then
        GBalAve = OldRound(GBalAve / GCustCnt)
      Else
        GBalAve = 0
      End If
    
      GFeeAve# = GrandFees#
      If GCustCnt > 0 Then
        GFeeAve# = OldRound(GFeeAve / GCustCnt)
      Else
        GFeeAve# = 0
      End If
      
      '                            18              19              20             21
      Print #RptHandle, AveBal(WhereFlag); dlm; AveBal(1); dlm; AveBal(2); dlm; AveBal(3); dlm;
      '                            22              23              24             25
      Print #RptHandle, AveFee(WhereFlag); dlm; AveFee(1); dlm; AveFee(2); dlm; AveFee(3); dlm;
      '                    26             27             28              29              30
      Print #RptHandle, GBalAve#; dlm; GFeeAve#; dlm; GCustCnt; dlm; GrandBal#; dlm; GrandFees#; dlm;
      If QPTrim$(fptxtCatCode.Text) = "ALL" Then
        '                         31                                 32                  33
        Print #RptHandle, CustCnt(WhereFlag); dlm; QPTrim$(fptxtCatCode.Text); dlm; TownRec.IssFee
      Else
        '                         31                                 32                                    33
        Print #RptHandle, CustCnt(WhereFlag); dlm; QPTrim$(fptxtCatCode.Text) + "/" + ThisCode; dlm; TownRec.IssFee
      End If

Inactive:
    frmBLShowPctComp.ShowPctComp PctCnt, NumOfCustIdxRecs * 3
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
    PctCnt = PctCnt + 1
    Next ZCnt&
    If CustCnt(WhereFlag) = 0 Then
      '                     0                1                2
      Print #RptHandle, CustCnt(1); dlm; CustCnt(2); dlm; CustCnt(3); dlm;
      '                     3                4                5
      Print #RptHandle, GTotBal(1); dlm; GTotBal(2); dlm; GTotBal(3); dlm;
      '                     6                  7                8
      Print #RptHandle, GTotFees(1); dlm; GTotFees(2); dlm; GTotFees(3); dlm;
      '                      9                     10                     11
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      '                             12                         13                 14
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      '                            15                       16
      Print #RptHandle, QPTrim$(TownRec.TownName); dlm; WhereFlag; dlm;
      If WhereFlag = 1 Then
        '                                17
        Print #RptHandle, "Business License: Inside City Limits"; dlm;
      ElseIf WhereFlag = 2 Then
        '                                17
        Print #RptHandle, "Business License: Outside City Limits"; dlm;
      Else
        '                                17
        Print #RptHandle, "Business License: No Location Saved"; dlm;
      End If
      '                            18              19              20             21
      Print #RptHandle, AveBal(WhereFlag); dlm; AveBal(1); dlm; AveBal(2); dlm; AveBal(3); dlm;
      '                            22              23              24             25
      Print #RptHandle, AveFee(WhereFlag); dlm; AveFee(1); dlm; AveFee(2); dlm; AveFee(3); dlm;
      '                    26             27             28              29              30
      Print #RptHandle, GBalAve#; dlm; GFeeAve#; dlm; GCustCnt; dlm; GrandBal#; dlm; GrandFees#; dlm;
      If QPTrim$(fptxtCatCode.Text) = "ALL" Then
        '                         31                                 32
        Print #RptHandle, CustCnt(WhereFlag); dlm; QPTrim$(fptxtCatCode.Text); dlm; TownRec.IssFee
      Else
        '                         31                                 32
        Print #RptHandle, CustCnt(WhereFlag); dlm; QPTrim$(fptxtCatCode.Text) + "/" + ThisCode; dlm; TownRec.IssFee
      End If
    End If
    WhereFlag = WhereFlag + 1
    If WhereFlag = 4 Then Exit Do
    LoopCnt = LoopCnt + 1
  Loop
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True

  Close         'Close all open files now
  
  If CustCnt(1) + CustCnt(2) + CustCnt(3) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers in category " + TCat$ + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  arBLInOut.Show
  frmBLLoadReport.Show
  
  MainLog ("'Inside/Outside City Limits Report' processed in graphics format.")
  Exit Sub
  
GetCustFee:
  
  CustFee# = 0
  FeeAmt1# = 0
  FeeAmt2# = 0
  FeeAmt3# = 0
  FeeAmt4# = 0
  FeeAmt5# = 0
  
  Prorate# = CustRec.Prorate
  
  If Prorate# >= 100 Or Prorate# < 0 Then
    Prorate# = 1
  Else
    Prorate# = OldRound#(Prorate# * 0.01)
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        CustRec.DESC1 = CodeRec.CODEDESC           'Reset Code Descriptions
        If CodeRec.CodeType = "F" Then
          FeeAmt1# = CodeRec.Fee
          FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV1
          FeeAmt1# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
          GoTo C2
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV1
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt1# < CodeRec.BaseAmt1 Then FeeAmt1# = CodeRec.BaseAmt1
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt1# < CodeRec.BaseAmt2 Then FeeAmt1# = CodeRec.BaseAmt2
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt1# < CodeRec.BaseAmt3 Then FeeAmt1# = CodeRec.BaseAmt3
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt1# < CodeRec.BaseAmt4 Then FeeAmt1# = CodeRec.BaseAmt4
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt1# < CodeRec.BaseAmt5 Then FeeAmt1# = CodeRec.BaseAmt5
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt1# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt1# < CodeRec.BaseAmt6 Then FeeAmt1# = CodeRec.BaseAmt6
            FeeAmt1# = OldRound(FeeAmt1# * Prorate#)
            GoTo C2
          End If
          
        End If
      End If
    Next Snt&
  End If
  
  
C2:
  
  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt2# = CodeRec.Fee
          FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV2
          FeeAmt2# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
          GoTo C3
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV2
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt2# < CodeRec.BaseAmt1 Then FeeAmt2# = CodeRec.BaseAmt1
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt2# < CodeRec.BaseAmt2 Then FeeAmt2# = CodeRec.BaseAmt2
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt2# < CodeRec.BaseAmt3 Then FeeAmt2# = CodeRec.BaseAmt3
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt2# < CodeRec.BaseAmt4 Then FeeAmt2# = CodeRec.BaseAmt4
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt2# < CodeRec.BaseAmt5 Then FeeAmt2# = CodeRec.BaseAmt5
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt2# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt2# < CodeRec.BaseAmt6 Then FeeAmt2# = CodeRec.BaseAmt6
            FeeAmt2# = OldRound(FeeAmt2# * Prorate#)
            GoTo C3
          End If
        End If
      End If
    Next Snt&
  End If
  
  
C3:
'  CustFee# = OldRound#(CustFee# + FeeAmt2#)
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt3# = CodeRec.Fee
          FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV3
          FeeAmt3# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
          GoTo c4
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV3
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt3# < CodeRec.BaseAmt1 Then FeeAmt3# = CodeRec.BaseAmt1
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt3# < CodeRec.BaseAmt2 Then FeeAmt3# = CodeRec.BaseAmt2
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt3# < CodeRec.BaseAmt3 Then FeeAmt3# = CodeRec.BaseAmt3
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt3# < CodeRec.BaseAmt4 Then FeeAmt3# = CodeRec.BaseAmt4
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt3# < CodeRec.BaseAmt5 Then FeeAmt3# = CodeRec.BaseAmt5
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
            GoTo c4
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt3# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt3# < CodeRec.BaseAmt6 Then FeeAmt3# = CodeRec.BaseAmt6
            FeeAmt3# = OldRound(FeeAmt3# * Prorate#)
           GoTo c4
          End If
        End If
      End If
    Next Snt&
  End If
  
c4:
'  CustFee# = OldRound#(CustFee# + FeeAmt3#)
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt4# = CodeRec.Fee
          FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV4
          FeeAmt4# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
          GoTo c5
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV4
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt4# < CodeRec.BaseAmt1 Then FeeAmt4# = CodeRec.BaseAmt1
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt4# < CodeRec.BaseAmt2 Then FeeAmt4# = CodeRec.BaseAmt2
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt4# < CodeRec.BaseAmt3 Then FeeAmt4# = CodeRec.BaseAmt3
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt4# < CodeRec.BaseAmt4 Then FeeAmt4# = CodeRec.BaseAmt4
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt4# < CodeRec.BaseAmt5 Then FeeAmt4# = CodeRec.BaseAmt5
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt4# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt4# < CodeRec.BaseAmt6 Then FeeAmt4# = CodeRec.BaseAmt6
            FeeAmt4# = OldRound(FeeAmt4# * Prorate#)
            GoTo c5
          End If
          
        End If
      End If
    Next Snt&
  End If
  
c5:
'  CustFee# = OldRound#(CustFee# + FeeAmt4#)
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        If CodeRec.CodeType = "F" Then
          FeeAmt5# = CodeRec.Fee
          FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV5
          FeeAmt5# = OldRound#(Mult * CodeRec.Fee)
          FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
          GoTo SkipEm
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV5
          If Revenue# <= CodeRec.Recpt1 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
            If FeeAmt5# < CodeRec.BaseAmt1 Then FeeAmt5# = CodeRec.BaseAmt1
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt2 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
            If FeeAmt5# < CodeRec.BaseAmt2 Then FeeAmt5# = CodeRec.BaseAmt2
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt3 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
            If FeeAmt5# < CodeRec.BaseAmt3 Then FeeAmt5# = CodeRec.BaseAmt3
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt4 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
            If FeeAmt5# < CodeRec.BaseAmt4 Then FeeAmt5# = CodeRec.BaseAmt4
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt5 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
            If FeeAmt5# < CodeRec.BaseAmt5 Then FeeAmt5# = CodeRec.BaseAmt5
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          If Revenue# <= CodeRec.Recpt6 Then
            FeeAmt5# = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
            If FeeAmt5# < CodeRec.BaseAmt6 Then FeeAmt5# = CodeRec.BaseAmt6
            FeeAmt5# = OldRound(FeeAmt5# * Prorate#)
            GoTo SkipEm
          End If
          
        End If
      End If
    Next Snt&
  End If
  
SkipEm:
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + TownRec.IssFee)
  
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicListRpt", "PrintText", Erl)
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

Private Sub fptxtCatCode_LostFocus()
  If Not IsNumeric(fptxtCatCode.Text) Then
    fptxtCatCode.Text = "ALL"
  End If

End Sub
