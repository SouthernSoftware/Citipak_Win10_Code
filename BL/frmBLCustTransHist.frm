VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustTransHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Transaction History"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCustTransHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6516
      Left            =   1920
      TabIndex        =   6
      Top             =   1174
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11493
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLCustTransHist.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   3030
         TabIndex        =   5
         Tag             =   $"frmBLCustTransHist.frx":08E6
         Top             =   4650
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
         ColDesigner     =   "frmBLCustTransHist.frx":099F
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2970
         TabIndex        =   0
         Tag             =   $"frmBLCustTransHist.frx":0C96
         Top             =   1590
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
         ColDesigner     =   "frmBLCustTransHist.frx":0D42
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   3216
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Business License Reports' menu."
         Top             =   5424
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
         ButtonDesigner  =   "frmBLCustTransHist.frx":1039
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   624
         Left            =   5436
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLCustTransHist.frx":1217
         Top             =   5436
         Width           =   1872
         _Version        =   131072
         _ExtentX        =   3302
         _ExtentY        =   1101
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
         ButtonDesigner  =   "frmBLCustTransHist.frx":12C2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCustList 
         Height          =   1125
         Left            =   5910
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   $"frmBLCustTransHist.frx":14A1
         Top             =   2205
         Width           =   1020
         _Version        =   131072
         _ExtentX        =   1799
         _ExtentY        =   1984
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
         ButtonDesigner  =   "frmBLCustTransHist.frx":1560
      End
      Begin EditLib.fpText fptxtFirst 
         Height          =   396
         Left            =   3456
         TabIndex        =   1
         Tag             =   $"frmBLCustTransHist.frx":1744
         Top             =   2208
         Width           =   2268
         _Version        =   196608
         _ExtentX        =   4000
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
      Begin EditLib.fpText fptxtLast 
         Height          =   396
         Left            =   3456
         TabIndex        =   2
         Tag             =   $"frmBLCustTransHist.frx":1851
         Top             =   2880
         Width           =   2268
         _Version        =   196608
         _ExtentX        =   4000
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
      Begin EditLib.fpDateTime fptxtBegin 
         Height          =   348
         Left            =   4176
         TabIndex        =   3
         Tag             =   "Enter the date the report will use as the beginning date for it's transaction search."
         Top             =   3504
         Width           =   1692
         _Version        =   196608
         _ExtentX        =   2984
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
         InvalidColor    =   12648447
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
         Text            =   "08/11/2003"
         DateCalcMethod  =   0
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
         ThreeDFrameColor=   13684944
         Appearance      =   2
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
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   348
         Left            =   4176
         TabIndex        =   4
         Tag             =   "Enter the last date the report will look for as it searches through the transaction records for data for this report."
         Top             =   4080
         Width           =   1692
         _Version        =   196608
         _ExtentX        =   2984
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
         InvalidColor    =   12648447
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
         Text            =   "08/11/2003"
         DateCalcMethod  =   0
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
         Appearance      =   2
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
      Begin fpBtnAtlLibCtl.fpBtn fpcmdHelp 
         Height          =   636
         Left            =   720
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   $"frmBLCustTransHist.frx":195B
         Top             =   5424
         Width           =   2172
         _Version        =   131072
         _ExtentX        =   3831
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
         ButtonDesigner  =   "frmBLCustTransHist.frx":1A2B
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
         Left            =   768
         TabIndex        =   18
         Top             =   6096
         Width           =   2100
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Date:"
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
         Left            =   1632
         TabIndex        =   16
         Top             =   4128
         Width           =   2268
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Date:"
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
         Left            =   1632
         TabIndex        =   15
         Top             =   3552
         Width           =   2268
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Customer Num:"
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
         Left            =   1008
         TabIndex        =   14
         Top             =   2976
         Width           =   2268
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
         Left            =   1344
         TabIndex        =   13
         Top             =   4752
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3948
         Left            =   624
         Top             =   1344
         Width           =   6780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First Customer Num:"
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
         Left            =   912
         TabIndex        =   12
         Top             =   2304
         Width           =   2364
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
         Left            =   1488
         TabIndex        =   11
         Top             =   1728
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Transaction History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   396
         Left            =   1728
         TabIndex        =   10
         Top             =   576
         Width           =   4524
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
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   480
      TabIndex        =   19
      Top             =   6720
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
      Top             =   1042
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLCustTransHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdCustList_Click()
  frmBLCustHistList.Show vbModal
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
      Call cmdCustList_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustTransHist.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim CHandle As Integer
  Dim CRec As ARCustRecType
  Dim NumOfCRecs As Integer
  
  lblBalloon.Visible = False
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  cmdCustList.ToolTipText = "Press to bring up an interactive customer list."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  fptxtFirst.ToolTipText = "Enter the customer for whom the report will begin."
'  fptxtLast.ToolTipText = "Enter the last customer for whom the report will end."
'  fptxtBegin.ToolTipText = "The report will get data for all transactions on or after this date."
'  fptxtEnd.ToolTipText = "The report will get data for all transactions on or before this date."
'  fpcmdHelp.ToolTipText = "Press to activate or deactivate instructional balloons."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdProcess.ToolTipText = "Press to activate this report."
  fptxtBegin = "01/01/" + Mid(Date, 7, 4)
  fptxtEnd = Date
  OpenCustFile CHandle
  NumOfCRecs = LOF(CHandle) / Len(CRec)
  If NumOfCRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  Get CHandle, 1, CRec
  fptxtFirst.Text = QPTrim$(CRec.CustNumb)
  Get CHandle, NumOfCRecs, CRec
  fptxtLast.Text = QPTrim$(CRec.CustNumb)
  Close CHandle
  fpcmbPrintOrder.Text = "Billing Name"
  fpcmbPrintOrder.AddItem "Billing Name"
  fpcmbPrintOrder.AddItem "Account Number"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"

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
      fptxtFirst.SetFocus
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
  DoEvents
  Unload frmBLCustTransHist
End Sub

Private Sub cmdProcess_Click()
  Dim First$
  Dim Last$
  
  On Error GoTo ERRORSTUFF
  
  First = QPTrim$(fptxtFirst.Text)
  Last = QPTrim$(fptxtLast.Text)
  If QPTrim$(fptxtFirst.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "ERROR: Please enter a valid customer number in the first customer field."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtFirst.SetFocus
    Close
    Exit Sub
  End If
  
  If QPTrim$(fptxtLast.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "ERROR: Please enter a valid customer number in the last customer field."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtLast.SetFocus
    Close
    Exit Sub
  End If
  
  If Val(fptxtFirst.Text) > Val(fptxtLast.Text) Then
    frmBLMessageBoxJr.Label1.Caption = "ERROR: The Customer Numbers entered should begin with the smallest and end with the largest."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFirst.SetFocus
    Exit Sub
  End If
  
  If Check4ValidCustNum(First$, "First") = False Then
    Exit Sub
  End If
  
  If Check4ValidCustNum(Last$, "Last") = False Then
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
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustTransHist", "cmdProcess_Click", Erl)
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

Private Function Check4ValidCustNum(ByVal CustNum$, ByVal WhichOne$) As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCRecs As Integer
  Dim x As Integer
  
  'this function catches errant customer numbers entered
  'manually by the user
  On Error GoTo ERRORSTUFF
  
  Check4ValidCustNum = True 'all OK at this point
  OpenCustFile CHandle
  NumOfCRecs = LOF(CHandle) / Len(CustRec)
  If NumOfCRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "ERROR: There are no customer records saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Check4ValidCustNum = False
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, CustRec
      If CustNum = QPTrim$(CustRec.CustNumb) Then
        Exit For
      End If
  Next x
  Close CHandle

  If x > NumOfCRecs Then
    frmBLMessageBoxJr.Label1.Caption = "ERROR: The " + WhichOne + " Customer Number entered is not valid. Please refer to the customer list for help."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    If WhichOne = "First" Then
      fptxtFirst.SetFocus
    Else
      fptxtLast.SetFocus
    End If
    Check4ValidCustNum = False
  End If
  
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustTransHist", "Check4ValidCustNum", Erl)
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
  

End Function

Private Sub PrintGraphics()
  Dim RunBalance#
  Dim TransRec As ARTransRecType
  Dim ARTransFile As Integer
  Dim ReportFile$
  Dim SubReportFile$
  Dim RptHandle As Integer
  Dim SubRptHandle As Integer
  Dim TransRecd&
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim TownName$
  Dim dlm$, x As Integer
  Dim CustName$
  Dim CNameIdxRec As CustNameIdxType
  Dim CNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim FirstDate As Integer
  Dim LastDate As Integer
  Dim CustNum$, ThisCnt As Integer
  Dim CustBal As Double
  Dim TransCnt As Long, y As Long
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName$ = QPTrim$(TownRec.TownName)
  
  RunBalance# = 0
  OpenTransFile ARTransFile

  ReportFile$ = "BLRPTS\ARTRHIST.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name" Then
    OpenCustNameIdxFile IdxHandle
    NumOfIdxRecs = LOF(IdxHandle) \ Len(CNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfIdxRecs = LOF(IdxHandle) \ Len(CNumIdxRec)
  End If
  
  If NumOfIdxRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "No customers saved in data base."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer
  ReDim ThisTransAmt(1 To 18) As Double
  ReDim ThisTransCnt(1 To 18) As Double
  ReDim ThisTransDsc(1 To 18) As String
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name" Then
    For x = 1 To NumOfIdxRecs
      Get IdxHandle, x, CNameIdxRec
      IdxRec(x) = CNameIdxRec.CustRec
    Next x
  Else
    For x = 1 To NumOfIdxRecs
      Get IdxHandle, x, CNumIdxRec
      IdxRec(x) = CNumIdxRec.CustRec
    Next x
  End If
  
  Close IdxHandle
  FirstDate = Date2Num(fptxtBegin.Text)
  LastDate = Date2Num(fptxtEnd.Text)
  OpenCustFile CHandle
  frmBLShowPctComp.Label1 = "Processing Transaction History"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  ReDim TransArray(1 To 1) As Long
  
  For x = 1 To NumOfIdxRecs
    Get CHandle, IdxRec(x), CustRec
    If Val(CustRec.CustNumb) >= Val(fptxtFirst.Text) And Val(CustRec.CustNumb) <= Val(fptxtLast.Text) Then
      CustName$ = QPTrim$(CustRec.BillName)
      CustNum$ = QPTrim$(CustRec.CustNumb)
      CustBal = CustRec.AcctBal
      TransRecd& = CustRec.FirstTrans
      TransCnt = 0
      'need this to print out starting with the latest date...since the transactions
      'are saved such that the oldest date comes first, code hasd to be arranged
      'to flip all valid transaction records
      Do While TransRecd& > 0
        Get ARTransFile, TransRecd&, TransRec
          If TransRec.TransDate >= FirstDate And TransRec.TransDate <= LastDate Then
            TransCnt = TransCnt + 1
            ReDim Preserve TransArray(1 To TransCnt) As Long
            TransArray(TransCnt) = TransRecd&
          End If
          TransRecd& = TransRec.NextTrans
      Loop
      If TransCnt = 0 Then GoTo NoTransHere
      y = TransCnt
      Do
        Get ARTransFile, TransArray(y), TransRec
          'at this point we have a qualifying transaction
          '                   0
          Print #RptHandle, TownName$; dlm;
          '                   1
          Print #RptHandle, MakeRegDate(TransRec.TransDate); dlm;
          '                   2
          Print #RptHandle, TransRec.TransDesc; dlm;
          'now determine the transaction type
          'if detailtranstype = 0 then this is either a transaction that took
          'place before this version or there was no defineable charge
          '(ie...only as issuance fee was charged...issuance fee does not
          'have a transaction description just for itself)
          If TransRec.DetailTransType = 0 And TransRec.TransType > 0 Then
            Select Case TransRec.TransType
            Case 1
              '                   3
              Print #RptHandle, "Charge"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                     4
              Print #RptHandle, TransRec.TransAmount; dlm;
              'the next 3 variables are used to collect data
              'to be printed out in the report summary
              ThisTransAmt(1) = ThisTransAmt(1) + TransRec.TransAmount
              ThisTransCnt(1) = ThisTransCnt(1) + 1
              ThisTransDsc(1) = "Charge"
            Case 2
              '                   3
              Print #RptHandle, "Payment"; dlm;
              RunBalance# = RunBalance# - TransRec.TransAmount
              '                       4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(2) = ThisTransAmt(2) + TransRec.TransAmount
              ThisTransCnt(2) = ThisTransCnt(2) + 1
              ThisTransDsc(2) = "Payment"
            Case 6
              '                    3
              Print #RptHandle, "Penalty"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(3) = ThisTransAmt(3) + TransRec.TransAmount
              ThisTransCnt(3) = ThisTransCnt(3) + 1
              ThisTransDsc(3) = "Penalty"
            Case 9
              '                    4
              Print #RptHandle, "Beg Bal"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(4) = ThisTransAmt(4) + TransRec.TransAmount
              ThisTransCnt(4) = ThisTransCnt(4) + 1
              ThisTransDsc(4) = "Beg Bal"
            Case 100
              '                    3
              Print #RptHandle, "DN Adj."; dlm;
              RunBalance# = RunBalance# - TransRec.TransAmount
              '                        4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(5) = ThisTransAmt(5) + TransRec.TransAmount
              ThisTransCnt(5) = ThisTransCnt(5) + 1
              ThisTransDsc(5) = "DN Adj."
            Case 101
              '                    3
              Print #RptHandle, "UP Adj."; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(6) = ThisTransAmt(6) + TransRec.TransAmount
              ThisTransCnt(6) = ThisTransCnt(6) + 1
              ThisTransDsc(6) = "UP Adj."
            Case Else
              '                    3
              Print #RptHandle, "Unknown"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, 0; dlm;
              ThisTransAmt(7) = ThisTransAmt(7) + TransRec.TransAmount
              ThisTransCnt(7) = ThisTransCnt(7) + 1
              ThisTransDsc(7) = "Unknown"
            End Select
          ElseIf TransRec.DetailTransType > 0 Then 'these transactions
          'took place only after this version was installed
            Select Case TransRec.DetailTransType
            Case 110      'lic
              '                    3
              Print #RptHandle, "License Charge"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(8) = ThisTransAmt(8) + TransRec.TransAmount
              ThisTransCnt(8) = ThisTransCnt(8) + 1
              ThisTransDsc(8) = "License Charge"
            Case 101
              '                    3
              Print #RptHandle, "Penalty Charge"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(9) = ThisTransAmt(9) + TransRec.TransAmount
              ThisTransCnt(9) = ThisTransCnt(9) + 1
              ThisTransDsc(9) = "Penalty Charge"
            Case 211
              '                    3
              Print #RptHandle, "Paid License and Penalty"; dlm;
              RunBalance# = RunBalance# - TransRec.TransAmount
              '                    4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(10) = ThisTransAmt(10) + TransRec.TransAmount
              ThisTransCnt(10) = ThisTransCnt(10) + 1
              ThisTransDsc(10) = "Paid License and Penalty"
            Case 210
              '                    3
              Print #RptHandle, "Paid License"; dlm;
              RunBalance# = RunBalance# - TransRec.TransAmount
              '                    4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(11) = ThisTransAmt(11) + TransRec.TransAmount
              ThisTransCnt(11) = ThisTransCnt(11) + 1
              ThisTransDsc(11) = "Paid License"
            Case 201
              '                    3
              Print #RptHandle, "Paid Penalty"; dlm;
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(12) = ThisTransAmt(12) + TransRec.TransAmount
              ThisTransCnt(12) = ThisTransCnt(12) + 1
              ThisTransDsc(12) = "Paid Penalty"
            Case 311    'dn adj
              '                    3
              Print #RptHandle, "Adjust Down License and Penalty"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(13) = ThisTransAmt(13) + TransRec.TransAmount
              ThisTransCnt(13) = ThisTransCnt(13) + 1
              ThisTransDsc(13) = "Adjust Down License and Penalty"
            Case 310    'dn adj
              '                    3
              Print #RptHandle, "Adjust Down License"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(14) = ThisTransAmt(14) + TransRec.TransAmount
              ThisTransCnt(14) = ThisTransCnt(14) + 1
              ThisTransDsc(14) = "Adjust Down License"
            Case 301    'dn adj
              '                    3
              Print #RptHandle, "Adjust Down Penalty"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, -TransRec.TransAmount; dlm;
              ThisTransAmt(15) = ThisTransAmt(15) + TransRec.TransAmount
              ThisTransCnt(15) = ThisTransCnt(15) + 1
              ThisTransDsc(15) = "Adjust Down Penalty"
            Case 411    'up adj
              If TransRec.TransType = 13 Then
                '
                Print #RptHandle, "Down Pay Adjust"; dlm;
              Else
                '                    3
                Print #RptHandle, "Adjust Up License and Penalty"; dlm;
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(16) = ThisTransAmt(16) + TransRec.TransAmount
              ThisTransCnt(16) = ThisTransCnt(16) + 1
              ThisTransDsc(16) = "Adj Up Lic/Pen/Dwn Pay Adj"
           Case 410    'up adj
              If TransRec.TransType = 13 Then
                '                    3
                Print #RptHandle, "Down Pay Adjust"; dlm;
              Else
                '                    3
                Print #RptHandle, "Adjust Up License"; dlm;
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(17) = ThisTransAmt(17) + TransRec.TransAmount
              ThisTransCnt(17) = ThisTransCnt(17) + 1
              ThisTransDsc(17) = "Adj Up Lic/Down Pay Adjust"
            Case 401    'up adj
              If TransRec.TransType = 13 Then
                '                    3
                Print #RptHandle, "Down Pay Adjust"; dlm;
              Else
                '                    3
                Print #RptHandle, "Adjust Up Penalty"; dlm;
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, TransRec.TransAmount; dlm;
              ThisTransAmt(18) = ThisTransAmt(18) + TransRec.TransAmount
              ThisTransCnt(18) = ThisTransCnt(18) + 1
              ThisTransDsc(18) = "Adj Up Penalty/Dwn Pay Adj"
            Case Else
              '                    3
              Print #RptHandle, "Unknown"; dlm;
              RunBalance# = RunBalance# + TransRec.TransAmount
              '                    4
              Print #RptHandle, 0; dlm;
              ThisTransAmt(7) = ThisTransAmt(7) + TransRec.TransAmount
              ThisTransCnt(7) = ThisTransCnt(7) + 1
              ThisTransDsc(7) = "Unknown"
            End Select
          Else
            '                  4        5
            Print #RptHandle, ""; dlm; ""; dlm;
          End If
          '                              5
          Print #RptHandle, TransRec.BalanceAfterTrans; dlm;              'RunBalance#
          '                      6              7                8
          Print #RptHandle, CustName$; dlm; CustNum$; dlm; CustBal
          ThisCnt = ThisCnt + 1
        y = y - 1
        If y = 0 Then Exit Do
      Loop
    End If
    frmBLShowPctComp.ShowPctComp x, NumOfIdxRecs
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
NoTransHere:
  Next x
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If ThisCnt = 0 Then
    If QPTrim$(fptxtFirst.Text) <> QPTrim$(fptxtLast) Then
      frmBLMessageBoxJr.Label1.Caption = "For customer numbers " + QPTrim$(fptxtFirst.Text) + " through " + QPTrim$(fptxtLast) + " there are no transactions recorded between " + fptxtBegin + " and " + fptxtEnd + " ."
    Else
      frmBLMessageBoxJr.Label1.Caption = "For customer number " + QPTrim$(fptxtFirst.Text) + " there are no transactions recorded between " + fptxtBegin + " and " + fptxtEnd + " ."
    End If
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  Close         'Close all open files now
    
  GoSub GetTally
  arBLCustTransHist.Show
  frmBLLoadReport.Show
  
  MainLog ("'Customer Transaction History' report run for customers " + QPTrim$(fptxtFirst.Text) + " thru " + QPTrim$(fptxtLast.Text) + " from " + fptxtBegin + " to " + fptxtEnd + " in graphics format.")
  
  Exit Sub
  
GetTally:
  SubReportFile$ = "BLRPTS\ARSUBTRHIST.RPT"
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle
  For x = 1 To 18
    If ThisTransAmt(x) > 0 Then
      Print #SubRptHandle, QPTrim$(ThisTransDsc(x)); dlm; ThisTransCnt(x); dlm; ThisTransAmt(x)
    End If
  Next x
  Close SubRptHandle
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustTransHist", "PrintGraphics", Erl)
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
  Dim RunBalance#
  Dim Page As Integer
  Dim TransRec As ARTransRecType
  Dim ARTransFile As Integer
  Dim MaxLines As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TransRecd&
  Dim LineCnt As Integer
  Dim FF$, x As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim CNameIdxRec As CustNameIdxType
  Dim CNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim FirstDate As Integer
  Dim LastDate As Integer
  Dim CustNum$, ThisCnt As Integer
  Dim FirstNum$, LmtCnt As Integer
  Dim LastNum$, TopHead As Boolean
  Dim NumOfIdxRecs As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim TownName$
  Dim TownRecCnt As Integer
  Dim ThisTab As Integer
  Dim PrintHeaderFlag As Boolean
  Dim CustBal As Double
  Dim TransCnt As Long, y As Long
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  OpenTownFile TownHandle
  TownRecCnt = LOF(TownHandle) / Len(TownRec)
  If TownRecCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No Town Control files have been saved. Please go to the Town Setup screen and save town data before continuing with this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName)
  FF$ = Chr$(12)
  RunBalance# = 0
  Page = 0
  TopHead = True
  MaxLines = 55
  ReportFile$ = "ARTRHIST.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  LmtCnt = 0
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name" Then
    OpenCustNameIdxFile IdxHandle
    NumOfIdxRecs = LOF(IdxHandle) \ Len(CNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfIdxRecs = LOF(IdxHandle) \ Len(CNumIdxRec)
  End If
  
  If NumOfIdxRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "No customers saved in data base."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer
  ReDim ThisTransAmt(1 To 18) As Double
  ReDim ThisTransCnt(1 To 18) As Double
  ReDim ThisTransDsc(1 To 18) As String
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name" Then
    For x = 1 To NumOfIdxRecs
      Get IdxHandle, x, CNameIdxRec
      IdxRec(x) = CNameIdxRec.CustRec
    Next x
  Else
    For x = 1 To NumOfIdxRecs
      Get IdxHandle, x, CNumIdxRec
      IdxRec(x) = CNumIdxRec.CustRec
    Next x
  End If
  
  Close IdxHandle
  
  FirstDate = Date2Num(fptxtBegin.Text)
  LastDate = Date2Num(fptxtEnd.Text)
  FirstNum = QPTrim$(fptxtFirst.Text)
  LastNum = QPTrim$(fptxtLast.Text)
  
  GoSub PrintCustHeader
  OpenCustFile CHandle
  OpenTransFile ARTransFile
  frmBLShowPctComp.Label1 = "Processing Transaction History"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False
  
  PrintHeaderFlag = False
  
  ReDim TransArray(1 To 1) As Long

  For x = 1 To NumOfIdxRecs
    Get CHandle, IdxRec(x), CustRec
    If Val(CustRec.CustNumb) >= FirstNum And Val(CustRec.CustNumb) <= LastNum Then
'      GoSub PrintSubHeader
      TransRecd& = CustRec.FirstTrans
      TransCnt = 0
      'need this to print out starting with the latest date...since the transactions
      'are saved such that the oldest date comes first, code hasd to be arranged
      'to flip all valid transaction records
      Do While TransRecd& > 0
        Get ARTransFile, TransRecd&, TransRec
          If TransRec.TransDate >= FirstDate And TransRec.TransDate <= LastDate Then
            TransCnt = TransCnt + 1
            ReDim Preserve TransArray(1 To TransCnt) As Long
            TransArray(TransCnt) = TransRecd&
          End If
          TransRecd& = TransRec.NextTrans
      Loop
      If TransCnt = 0 Then GoTo NoTransHere
      y = TransCnt
       Do
         Get ARTransFile, TransArray(y), TransRec
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintCustHeader
            GoSub PrintSubHeader
          End If
          If PrintHeaderFlag = False Then
            GoSub PrintSubHeader
            PrintHeaderFlag = True
          End If
          Print #RptHandle, MakeRegDate(TransRec.TransDate);
          Print #RptHandle, Tab(14); QPTrim$(TransRec.TransDesc);
          
          If TransRec.DetailTransType = 0 And TransRec.TransType > 0 Then
            Select Case TransRec.TransType
            Case 1
              Print #RptHandle, Tab(29); "Charge";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(1) = ThisTransAmt(1) + TransRec.TransAmount
              ThisTransCnt(1) = ThisTransCnt(1) + 1
              ThisTransDsc(1) = "Charge"
            Case 2
              Print #RptHandle, Tab(29); "Payment";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(2) = ThisTransAmt(2) + TransRec.TransAmount
              ThisTransCnt(2) = ThisTransCnt(2) + 1
              ThisTransDsc(2) = "Payment"
            Case 6
              Print #RptHandle, Tab(29); "Penalty";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(3) = ThisTransAmt(3) + TransRec.TransAmount
              ThisTransCnt(3) = ThisTransCnt(3) + 1
              ThisTransDsc(3) = "Penalty"
            Case 9
              Print #RptHandle, Tab(29); "Beg Bal";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(4) = ThisTransAmt(4) + TransRec.TransAmount
              ThisTransCnt(4) = ThisTransCnt(4) + 1
              ThisTransDsc(4) = "Beg Bal"
            Case 100
              Print #RptHandle, Tab(29); "DN Adj.";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(5) = ThisTransAmt(5) + TransRec.TransAmount
              ThisTransCnt(5) = ThisTransCnt(5) + 1
              ThisTransDsc(5) = "DN Adj."
            Case 101
              Print #RptHandle, Tab(29); "UP Adj.";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(6) = ThisTransAmt(6) + TransRec.TransAmount
              ThisTransCnt(6) = ThisTransCnt(6) + 1
              ThisTransDsc(6) = "UP Adj."
            Case Else
              Print #RptHandle, Tab(29); "Unknown";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", 0);
              ThisTransAmt(7) = ThisTransAmt(7) + TransRec.TransAmount
              ThisTransCnt(7) = ThisTransCnt(7) + 1
              ThisTransDsc(7) = "Unknown"
            End Select
          ElseIf TransRec.DetailTransType > 0 Then
            Select Case TransRec.DetailTransType
            Case 110
              Print #RptHandle, Tab(29); "License Charge";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(8) = ThisTransAmt(8) + TransRec.TransAmount
              ThisTransCnt(8) = ThisTransCnt(8) + 1
              ThisTransDsc(8) = "License Charge"
            Case 101
              Print #RptHandle, Tab(29); "Penalty Charge";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(9) = ThisTransAmt(9) + TransRec.TransAmount
              ThisTransCnt(9) = ThisTransCnt(9) + 1
              ThisTransDsc(9) = "Penalty Charge"
            Case 211
              Print #RptHandle, Tab(29); "Paid License and Penalty";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(10) = ThisTransAmt(10) + TransRec.TransAmount
              ThisTransCnt(10) = ThisTransCnt(10) + 1
              ThisTransDsc(10) = "Paid License and Penalty"
            Case 210
              Print #RptHandle, Tab(29); "Paid License";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(11) = ThisTransAmt(11) + TransRec.TransAmount
              ThisTransCnt(11) = ThisTransCnt(11) + 1
              ThisTransDsc(11) = "Paid License"
            Case 201
              Print #RptHandle, Tab(29); "Paid Penalty";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(12) = ThisTransAmt(12) + TransRec.TransAmount
              ThisTransCnt(12) = ThisTransCnt(12) + 1
              ThisTransDsc(12) = "Paid Penalty"
            Case 311    'dn adj
              Print #RptHandle, Tab(29); "Adjust Down License and Penalty";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(13) = ThisTransAmt(13) + TransRec.TransAmount
              ThisTransCnt(13) = ThisTransCnt(13) + 1
              ThisTransDsc(13) = "Adjust Down License and Penalty"
            Case 310    'dn adj
              Print #RptHandle, Tab(29); "Adjust Down License";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(14) = ThisTransAmt(14) + TransRec.TransAmount
              ThisTransCnt(14) = ThisTransCnt(14) + 1
              ThisTransDsc(14) = "Adjust Down License"
            Case 301    'dn adj
              Print #RptHandle, Tab(29); "Adjust Down Penalty";
              RunBalance# = RunBalance# - TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", -TransRec.TransAmount);
              ThisTransAmt(15) = ThisTransAmt(15) + TransRec.TransAmount
              ThisTransCnt(15) = ThisTransCnt(15) + 1
              ThisTransDsc(15) = "Adjust Down Penalty"
            Case 411    'up adj
              If TransRec.TransType = 13 Then
                Print #RptHandle, Tab(29); "Down Pay Adjust";
              Else
                Print #RptHandle, Tab(29); "Adjust Up License and Penalty";
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(16) = ThisTransAmt(16) + TransRec.TransAmount
              ThisTransCnt(16) = ThisTransCnt(16) + 1
              ThisTransDsc(16) = "Adj Up Lic/Pen & Dwn Pay Adj"
            Case 410    'up adj
              If TransRec.TransType = 13 Then
                Print #RptHandle, Tab(29); "Down Pay Adjust";
              Else
                Print #RptHandle, Tab(29); "Adjust Up License";
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(17) = ThisTransAmt(17) + TransRec.TransAmount
              ThisTransCnt(17) = ThisTransCnt(17) + 1
              ThisTransDsc(17) = "Adj Up Lic/Dwn Pay Adj"
            Case 401    'up adj
              If TransRec.TransType = 13 Then
                Print #RptHandle, Tab(29); "Down Pay Adj";
              Else
                Print #RptHandle, Tab(29); "Adjust Up Penalty";
              End If
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", TransRec.TransAmount);
              ThisTransAmt(18) = ThisTransAmt(18) + TransRec.TransAmount
              ThisTransCnt(18) = ThisTransCnt(18) + 1
              ThisTransDsc(18) = "Adj Up Pen/Dwn Pay Adj"
            Case Else
              Print #RptHandle, Tab(29); "Unknown";
              RunBalance# = RunBalance# + TransRec.TransAmount
              Print #RptHandle, Tab(58); Using("###,##0.00", 0);
              ThisTransAmt(7) = ThisTransAmt(7) + TransRec.TransAmount
              ThisTransCnt(7) = ThisTransCnt(7) + 1
              ThisTransDsc(7) = "Unknown"
            End Select
          End If
        ThisCnt = ThisCnt + 1
        Print #RptHandle, Tab(71); Using("###,##0.00", TransRec.BalanceAfterTrans)               'RunBalance#
        LmtCnt = LmtCnt + 1
        LineCnt = LineCnt + 1
NoTrans:
        y = y - 1
        If y = 0 Then Exit Do
      Loop
      If LmtCnt > 0 Then
        Print #RptHandle, ""
        Print #RptHandle, "Number Of Transactions: " + CStr(LmtCnt)
        Print #RptHandle, ""
        LmtCnt = 0
        LineCnt = LineCnt + 3
      End If
    End If
    If LineCnt >= MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintCustHeader
    End If
    frmBLShowPctComp.ShowPctComp x, NumOfIdxRecs
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
    PrintHeaderFlag = False
NoTransHere:
  Next x
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True
  
  If ThisCnt = 0 Then
    If QPTrim$(fptxtFirst.Text) <> QPTrim$(fptxtLast) Then
      frmBLMessageBoxJr.Label1.Caption = "For customer numbers " + QPTrim$(fptxtFirst.Text) + " through " + QPTrim$(fptxtLast) + " there are no transactions recorded between " + fptxtBegin + " and " + fptxtEnd + " ."
    Else
      frmBLMessageBoxJr.Label1.Caption = "For customer number " + QPTrim$(fptxtFirst.Text) + " there are no transactions recorded between " + fptxtBegin + " and " + fptxtEnd + " ."
    End If
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  Print #RptHandle, Chr$(18);
  Print #RptHandle, FF$;
  
  GoSub GetTally
  Close         'Close all open files now

  ViewPrint ReportFile$, "Customer Account History", True
  
  KillFile ReportFile$
  
  MainLog ("'Customer Transaction History' report run for customers " + QPTrim$(fptxtFirst.Text) + " thru " + QPTrim$(fptxtLast.Text) + " from " + fptxtBegin + " to " + fptxtEnd + " in text format.")
  Exit Sub

PrintCustHeader:
  ThisTab = Len(TownName$)
  ThisTab = 40 - (ThisTab / 2)
  
  TopHead = True
  Page = Page + 1
  Print #RptHandle, Tab(ThisTab); TownName$
  Print #RptHandle, Tab(17); "Business License :  Customer Transaction History"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "  Date"; Tab(12); "Description "; Tab(33); "Type"; Tab(60); "  Amount"; Tab(74); "Balance"
  Print #RptHandle, String$(80, "=")
  LineCnt = 4
  Return
  
PrintSubHeader:
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintCustHeader
  End If
  If TopHead = True Then
    Print #RptHandle, "Billing Name: "; QPTrim$(CustRec.BillName); "  Customer Number: "; QPTrim$(CustRec.CustNumb);
    Print #RptHandle, Tab(71); Using("###,##0.00", CustRec.AcctBal)
    Print #RptHandle, String$(80, "=")
    TopHead = False
    LineCnt = LineCnt + 2
  Else
    Print #RptHandle, ""
    Print #RptHandle, "Billing Name: "; QPTrim$(CustRec.BillName); "  Customer Number: "; QPTrim$(CustRec.CustNumb);
    Print #RptHandle, Tab(71); Using("###,##0.00", CustRec.AcctBal)
    Print #RptHandle, String$(80, "=")
    LineCnt = LineCnt + 3
  End If
  
  Return

GetTally:
  Print #RptHandle, FF$
  ThisTab = Len(TownName$)
  ThisTab = 40 - (ThisTab / 2)
  
  TopHead = True
  Page = Page + 1
  Print #RptHandle, Tab(ThisTab); TownName$
  Print #RptHandle, Tab(17); "Business License :  Customer Transaction History"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, Tab(5); "Transaction Description"; Tab(35); "Number Of Transactions"; Tab(63); "Transaction Amount"
  Print #RptHandle, String$(80, "=")
  For x = 1 To 18
    If ThisTransCnt(x) > 0 Then
      Print #RptHandle, Tab(5); QPTrim$(ThisTransDsc(x)); Tab(40); Using("###,###", ThisTransCnt(x)); Tab(66); Using("$###,###,##0.00", ThisTransAmt(x))
    End If
  Next x
  
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustTransHist", "PrintText", Erl)
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
    cmdCustList.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    fptxtFirst.ToolTipText = ""
    fptxtLast.ToolTipText = ""
    fptxtBegin.ToolTipText = ""
    fptxtEnd.ToolTipText = ""
    fpcmdHelp.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    cmdCustList.ToolTipText = "Press to bring up an interactive customer list."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    fptxtFirst.ToolTipText = "Enter the customer for whom the report will begin."
'    fptxtLast.ToolTipText = "Enter the last customer for whom the report will end."
'    fptxtBegin.ToolTipText = "The report will get data for all transactions on or after this date."
'    fptxtEnd.ToolTipText = "The report will get data for all transactions on or before this date."
'    fpcmdHelp.ToolTipText = "Press to activate or deactivate instructional balloons."
'    cmdExit.ToolTipText = "Press to exit this screen."
'    cmdProcess.ToolTipText = "Press to activate this report."
  End If

End Sub
