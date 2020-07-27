VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPrintLic 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Business License "
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLPrintLic.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7716
      Left            =   720
      TabIndex        =   15
      Top             =   480
      Width           =   10236
      _Version        =   196609
      _ExtentX        =   18055
      _ExtentY        =   13610
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLPrintLic.frx":08CA
      Begin LpLib.fpCombo fpcmbBalanceType 
         Height          =   405
         Left            =   6525
         TabIndex        =   11
         Tag             =   $"frmBLPrintLic.frx":08E6
         Top             =   5610
         Width           =   2790
         _Version        =   196608
         _ExtentX        =   4921
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
         ColDesigner     =   "frmBLPrintLic.frx":0AD6
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   6045
         TabIndex        =   9
         Tag             =   $"frmBLPrintLic.frx":0DD1
         Top             =   3510
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
         ColDesigner     =   "frmBLPrintLic.frx":0E7D
      End
      Begin LpLib.fpCombo fpcmbPrintFeesYN 
         Height          =   405
         Left            =   7440
         TabIndex        =   10
         Tag             =   $"frmBLPrintLic.frx":1178
         Top             =   4605
         Width           =   975
         _Version        =   196608
         _ExtentX        =   1720
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
         ColDesigner     =   "frmBLPrintLic.frx":132D
      End
      Begin EditLib.fpText fptxtBegNum 
         Height          =   396
         Left            =   1728
         TabIndex        =   2
         Tag             =   $"frmBLPrintLic.frx":1628
         Top             =   3408
         Width           =   1548
         _Version        =   196608
         _ExtentX        =   2730
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
         BackColor       =   16777215
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
         MaxLength       =   12
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   5616
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'License Processing' menu."
         Top             =   6624
         Width           =   1644
         _Version        =   131072
         _ExtentX        =   2900
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
         ButtonDesigner  =   "frmBLPrintLic.frx":1970
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   7536
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Press 'Process' to begin printing the business license forms using the parameters entered above."
         Top             =   6624
         Width           =   1644
         _Version        =   131072
         _ExtentX        =   2900
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
         ButtonDesigner  =   "frmBLPrintLic.frx":1B4F
      End
      Begin EditLib.fpDateTime fptxtVThru 
         Height          =   370
         Left            =   2400
         TabIndex        =   1
         Tag             =   "The date entered here will appear on the business license forms as the expiration date for this license."
         Top             =   2400
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   653
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpDateTime fptxtFromDate 
         Height          =   370
         Left            =   2400
         TabIndex        =   0
         Tag             =   "The date entered here will appear on the business license as the first day of the valid date range for this license."
         Top             =   1872
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   653
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpText fptxtHeading 
         Height          =   396
         Index           =   0
         Left            =   864
         TabIndex        =   3
         Tag             =   $"frmBLPrintLic.frx":1D2E
         Top             =   4464
         Width           =   4956
         _Version        =   196608
         _ExtentX        =   8742
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
         MaxLength       =   255
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
      Begin EditLib.fpText fptxtHeading 
         Height          =   396
         Index           =   1
         Left            =   864
         TabIndex        =   4
         Tag             =   $"frmBLPrintLic.frx":1E0B
         Top             =   4896
         Width           =   4956
         _Version        =   196608
         _ExtentX        =   8742
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
         MaxLength       =   255
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
      Begin EditLib.fpText fptxtHeading 
         Height          =   396
         Index           =   2
         Left            =   864
         TabIndex        =   5
         Tag             =   $"frmBLPrintLic.frx":1EE9
         Top             =   5328
         Width           =   4956
         _Version        =   196608
         _ExtentX        =   8742
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
         MaxLength       =   255
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
      Begin EditLib.fpText fptxtHeading 
         Height          =   396
         Index           =   3
         Left            =   864
         TabIndex        =   6
         Tag             =   $"frmBLPrintLic.frx":1FC6
         Top             =   5760
         Width           =   4956
         _Version        =   196608
         _ExtentX        =   8742
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
         MaxLength       =   255
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   636
         Left            =   3696
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintLic.frx":20A4
         Top             =   6624
         Width           =   1644
         _Version        =   131072
         _ExtentX        =   2900
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
         ButtonDesigner  =   "frmBLPrintLic.frx":2196
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   636
         Left            =   1248
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintLic.frx":2372
         Top             =   6624
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
         ButtonDesigner  =   "frmBLPrintLic.frx":2442
      End
      Begin EditLib.fpDateTime fptxtIssDate 
         Height          =   370
         Left            =   7776
         TabIndex        =   7
         Tag             =   "The date entered here will be recorded as the post date for this transaction. It does not appear on license forms."
         ToolTipText     =   "Enter the date which will appear on the business licenses indicating the first valid day."
         Top             =   1488
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   653
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpDateTime fpBLYear 
         Height          =   370
         Left            =   7056
         TabIndex        =   8
         Tag             =   "The date entered here will appear on the business license as the active year for this license"
         Top             =   2544
         Width           =   972
         _Version        =   196608
         _ExtentX        =   1714
         _ExtentY        =   653
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
         Text            =   "2018"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdList 
         Height          =   348
         Left            =   3312
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintLic.frx":2625
         Top             =   3438
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   614
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
         ButtonDesigner  =   "frmBLPrintLic.frx":2727
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Business License For Year:"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   2160
         Width           =   2988
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date License Issued:"
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
         Left            =   5280
         TabIndex        =   28
         Top             =   1536
         Width           =   2412
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
         Left            =   1296
         TabIndex        =   26
         Top             =   7296
         Width           =   2100
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   1728
         TabIndex        =   24
         Top             =   2448
         Width           =   492
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         TabIndex        =   23
         Top             =   1920
         Width           =   732
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balances To Print On License"
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
         Left            =   6192
         TabIndex        =   22
         Top             =   5280
         Width           =   3228
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print License Fees (Y/N)?"
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
         Left            =   6576
         TabIndex        =   21
         Top             =   4272
         Width           =   2796
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   5340
         Left            =   384
         Top             =   1152
         Width           =   9516
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "License Heading:"
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
         Left            =   2352
         TabIndex        =   20
         Top             =   4128
         Width           =   1980
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning License Number:"
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
         Left            =   1728
         TabIndex        =   19
         Top             =   3024
         Width           =   3036
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Business License Date Range:"
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
         Left            =   1344
         TabIndex        =   18
         Top             =   1536
         Width           =   3276
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   2832
         Top             =   288
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Form Fed Business License"
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
         Left            =   2976
         TabIndex        =   17
         Top             =   432
         Width           =   4572
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
         Left            =   7152
         TabIndex        =   16
         Top             =   3168
         Width           =   1308
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   864
      TabIndex        =   27
      Top             =   8496
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
      MaxWidth        =   6000
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
      Height          =   8124
      Left            =   504
      Top             =   276
      Width           =   10644
   End
End
Attribute VB_Name = "frmBLPrintLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim UsePermLicNum As Boolean
Private Sub cmdAlign_Click()
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim SHeading4$
  Dim Heading1 As Integer
  Dim Heading2 As Integer
  Dim Heading3 As Integer
  Dim Heading4 As Integer
  Dim tab1 As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim ReportFile$
  Dim LPRINT As Integer
  Dim LCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  SHeading1$ = QPTrim$(fptxtHeading(0).Text)
  SHeading2$ = QPTrim$(fptxtHeading(1).Text)
  SHeading3$ = QPTrim$(fptxtHeading(2).Text)
  SHeading4$ = QPTrim$(fptxtHeading(3).Text)

  Heading1 = Len(SHeading1$)
  Heading2 = Len(SHeading2$)
  Heading3 = Len(SHeading3$)
  Heading4 = Len(SHeading4$)

  If Len(Heading1) > 0 Then tab1 = Heading1 / 2 Else tab1 = 0
  If Len(Heading2) > 0 Then tab2 = Heading2 / 2 Else tab2 = 0
  If Len(Heading3) > 0 Then Tab3 = Heading3 / 2 Else Tab3 = 0
  If Len(Heading4) > 0 Then Tab4 = Heading4 / 2 Else Tab4 = 0
  
  ReportFile$ = "LICMASK.PRT"
  LPRINT = FreeFile
  Open ReportFile$ For Output As #LPRINT

  ' Print Form Test
  Print #LPRINT, "TOP"
  For LCnt = 1 To 4
    Print #LPRINT, ""
  Next LCnt
  Print #LPRINT, Tab(37 - tab1); SHeading1$
  Print #LPRINT, Tab(37 - tab2); SHeading2$
  Print #LPRINT, Tab(37 - Tab3); SHeading3$
  Print #LPRINT, Tab(37 - Tab4); SHeading4$
  Print #LPRINT, Tab(66); Mid(fptxtVThru.Text, 7, 4)
  Print #LPRINT,
  Print #LPRINT, Tab(11); "Name of Some Business"
  Print #LPRINT, Tab(11); "Address Line 1"; Tab(58); "########"
  Print #LPRINT, Tab(11); "Address Line 2"
  Print #LPRINT, Tab(11); "Address Line 3"
  Print #LPRINT, Tab(55); Mid(fptxtFromDate.Text, 1, 6) + Mid(fptxtFromDate.Text, 9, 2);
  Print #LPRINT, Tab(64); Mid(fptxtVThru.Text, 1, 6) + Mid(fptxtVThru.Text, 9, 2)
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(11); String$(35, "X")
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(5); "XXXXXXXX"; Tab(15); String$(30, "X"); Tab(62); "XXXXX.XX"
  For LCnt = 24 To 35
    Print #LPRINT, ""
  Next LCnt
  Print #LPRINT, Tab(62); "XXXXX.XX"
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(62); "XXXXX.XX"
  Print #LPRINT,
  Print #LPRINT, "BOTTOM"
  
  Close
  ViewPrint ReportFile, "Business License Mask", True
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLic", "cmdAlign_Click", Erl)
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

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtVThru.ToolTipText = ""
    fptxtFromDate.ToolTipText = ""
    fptxtBegNum.ToolTipText = ""
    cmdList.ToolTipText = ""
    fptxtHeading(0).ToolTipText = ""
    fptxtHeading(1).ToolTipText = ""
    fptxtHeading(2).ToolTipText = ""
    fptxtHeading(3).ToolTipText = ""
'    fpcmbFeeYN.ToolTipText = ""
    fpcmbPrintFeesYN.ToolTipText = ""
    fpcmbBalanceType.ToolTipText = ""
    cmdAlign.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    fpcmbPrintOrder.ToolTipText = ""
    cmdHelp.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtVThru.ToolTipText = "The date entered here will appear on the license as the expiration date."
'    fptxtIssueDate.ToolTipText = "Enter the date which will appear on the business licenses the first valid day of this license period."
'    fptxtBegNum.ToolTipText = "Enter the new business license number that will begin the license printing process."
'    cmdList.ToolTipText = "Use this button to bring up a list of all customer's and their license numbers."
'    fptxtHeading(0).ToolTipText = "Optional line of text that will appear as the first line of the license header."
'    fptxtHeading(1).ToolTipText = "Optional line of text that will appear as the second line of the license header."
'    fptxtHeading(2).ToolTipText = "Optional line of text that will appear as the third line of the license header."
'    fptxtHeading(3).ToolTipText = "Optional line of text that will appear as the fourth line of the license header."
'    fpcmbFeeYN.ToolTipText = "Select Yes and any fees calculated and not posted for these licenses will be reset to zero. Choose No to allow the unposted calculations to remain."
'    fpcmbPrintFeesYN.ToolTipText = "This option allows the current fees to appear on each license. This option is disabled if the 'Charge Account With Fee Y/N' option is No."
'    fpcmbBalanceType.ToolTipText = "Business licenses can be printed with total outstanding balances and current balances or just current balances."
'    cmdAlign.ToolTipText = "Use this button to help line up license forms."
'    cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'    cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  End If
'  frmBLMessageBox.Label1.Caption = "If you select No in the 'Charge Account With Fee (Y/N)?' option box then all unposted calculations are reset to zero. To reverse this option once licenses are printed you must re-process the register. If this field is set to 'N' then 'Print License Fees Y/N' and 'Balance To Print On License' are disabled."
'  frmBLMessageBox.Label1.Height = 1300
'  frmBLMessageBox.Label1.Top = 300
'  frmBLMessageBox.Label2.Caption = "The 'Print License Fees (Y/N)?' option gives you the ability to print business licenses with no fees appearing on the license itself. This has no bearing on the fee calculations. When posting takes place each customer will be assessed the license fees calculated for them."
'  frmBLMessageBox.Label2.Top = 1800
'  frmBLMessageBox.Label2.Height = 1000
'  frmBLMessageBox.Label3.Caption = "The 'Balances To Print On License' option allows licenses to be printed with current as well as total outstanding balances or just current balances."
'  frmBLMessageBox.Label3.Height = 1500
'  frmBLMessageBox.Label3.Top = 3100
'  frmBLMessageBox.Show vbModal
End Sub

Private Sub cmdList_Click()
  frmBLLicenseNumList.Show vbModal
End Sub

Private Sub cmdProcess_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    fpcmbPrintOrder.SetFocus
  End If
End Sub

Private Sub fpcmbBalanceType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBalanceType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBalanceType.ListIndex = -1
  End If
  If fpcmbBalanceType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtFromDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

'Private Sub fpcmbFeeYN_Change()
'  If QPTrim$(fpcmbFeeYN.Text) = "" Then
'    fpcmbFeeYN.Text = "Yes"
'  End If
'  If QPTrim$(fpcmbFeeYN.Text) = "Yes" Then
'    fpcmbPrintFeesYN.Enabled = True
'    fpcmbBalanceType.Enabled = True
'  Else
'    fpcmbPrintFeesYN.Enabled = False
'    fpcmbBalanceType.Enabled = False
'  End If
'End Sub

'Private Sub fpcmbFeeYN_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcmbFeeYN.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbFeeYN.ListIndex = -1
'  End If
'  If fpcmbFeeYN.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fpcmbPrintFeesYN.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub

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
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintLic.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbPrintFeesYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintFeesYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintFeesYN.ListIndex = -1
  End If
  If fpcmbPrintFeesYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbBalanceType.Enabled = True Then
        fpcmbBalanceType.SetFocus
      Else
        fptxtFromDate.SetFocus
      End If
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
      fpcmbPrintFeesYN.SetFocus
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
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLPrintLic
End Sub

Private Sub cmdProcess_Click()

  On Error Resume Next
  If Not Exist("artmppst.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please process Business License registers first."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(fptxtBegNum.Text) = "" Then
    fptxtBegNum.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please enter a value for license number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtBegNum.BackColor = &HFFFFFF
    fptxtBegNum.SetFocus
    Exit Sub
  End If
  
  If Date2Num(fptxtVThru.Text) < Date2Num(fptxtFromDate.Text) Then
    frmBLMessageBoxJr.Label1.Caption = "The 'From' date comes after the new expiration date. Please revise these dates."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFromDate.SetFocus
    Exit Sub
  ElseIf Date2Num(fptxtVThru.Text) = Date2Num(fptxtFromDate.Text) Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The 'From' date and the new expiration date are the same. If this is correct then press F10 to continue. Otherwise press ESC to return to the screen."
    frmBLMessageBoxJrWOpts.Label1.Top = 600
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      fptxtVThru.SetFocus
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  If QPTrim$(fptxtBegNum.Text) <> "PERMANENT" Then
    If Look4DupLicNums = True Then
      Unload frmBLLoadReport
      Exit Sub
    End If
  End If
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  Call PrintText
  
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double, y As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim StoreExpireDate$
  Dim ExpireDate$
  Dim Year$
  Dim NumOfTransRecs As Double
  Dim NextTransRec As Double
  Dim CategoryRecord1 As Integer
  Dim CategoryRecord2 As Integer
  Dim CategoryRecord3 As Integer
  Dim CategoryRecord4 As Integer
  Dim CategoryRecord5 As Integer
  Dim TotalBillAmt#
  Dim PostDate$
  Dim CustomerNumber As Integer
  Dim Prev As Long
  Dim CategoryDesc$
  Dim CategoryDesc1$
  Dim CategoryDesc2$
  Dim CategoryDesc3$
  Dim CategoryDesc4$
  Dim CategoryDesc5$, DidCnt As Integer
  Dim LICENSE#, ll As Integer
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim Heading1 As Integer
  Dim Heading2 As Integer
  Dim Heading3 As Integer
  Dim Heading4 As Integer
  Dim tab1 As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim SHeading4$
  Dim FromDate$
  Dim SCnt As Integer, LCnt As Integer
  Dim TempHandle As Integer
  Dim TempRec As TempTransPostType
  Dim TempNum As Integer
  Dim TPHandle As Integer
  Dim TempLPRec As TempLicPrintType
  Dim TempPrintNum As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim SeqNum As Integer
  Dim NumOfTempRecs As Integer
  Dim PrintFees As Boolean
  Dim IssFee As Double
  Dim BalanceFlag As Integer
  Dim Nextx As Double
  Dim ThisDate$
  Dim ThisLen As Integer
  Dim DHandle As Integer
  Dim ThisHeader$
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
    IssFee = TownRec.IssFee
  Else
    IssFee = 0
  End If
  
  SeqNum = 1
  If UsePermLicNum = False Then
    If QPTrim$(fptxtBegNum.Text) = "" Then
      fptxtBegNum.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "Please enter a beginning license number."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      fptxtBegNum.BackColor = &HFFFFFF
      fptxtBegNum.SetFocus
      Close
      Exit Sub
    End If
    LICENSE# = QPTrim$(fptxtBegNum.Text)
  End If
  
  SHeading1$ = QPTrim$(fptxtHeading(0).Text)
  SHeading2$ = QPTrim$(fptxtHeading(1).Text)
  SHeading3$ = QPTrim$(fptxtHeading(2).Text)
  SHeading4$ = QPTrim$(fptxtHeading(3).Text)

  Heading1 = Len(SHeading1$)
  Heading2 = Len(SHeading2$)
  Heading3 = Len(SHeading3$)
  Heading4 = Len(SHeading4$)

  If Len(Heading1) > 0 Then tab1 = Heading1 / 2 Else tab1 = 0
  If Len(Heading2) > 0 Then tab2 = Heading2 / 2 Else tab2 = 0
  If Len(Heading3) > 0 Then Tab3 = Heading3 / 2 Else Tab3 = 0
  If Len(Heading4) > 0 Then Tab4 = Heading4 / 2 Else Tab4 = 0
  
  StoreExpireDate$ = fptxtVThru.Text
  ExpireDate$ = Mid(fptxtVThru.Text, 1, 6) + Mid(fptxtVThru.Text, 9, 2)
  Year$ = Mid(fptxtVThru.Text, 7, 4)
  
  FromDate$ = Mid(fptxtFromDate.Text, 1, 6) + Mid(fptxtFromDate.Text, 9, 2)
  PostDate$ = fptxtIssDate.Text
  ReportFile$ = "ARLICLST.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  PrintFees = False
  If QPTrim$(fpcmbPrintFeesYN.Text) = "Yes" Then
    PrintFees = True
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  ReDim IdxRecs(1 To 1) As Double
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
    TempLPRec.Order = "A"
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
    TempLPRec.Order = "B"
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
  
  Close IdxHandle
  
  OpenCustFile CustHandle
  
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  Close THandle
  NextTransRec = NumOfTransRecs + 1
  ' Print Main Body
  
  TempNum = 1
  If Not Exist("artmppst.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please process Business License registers first."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenTempPostFile TempHandle 'data posted from this
  NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
  ReDim PrintIdx(1 To 1) As Double

  Nextx = 0
  frmBLShowPctComp.Label1 = "Gathering Customer Data"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdAlign.Enabled = False
  cmdHelp.Enabled = False
  frmBLShowPctComp.cmdCancel.Visible = False
  
  For x = 1 To NumOfCustIdxRecs
    For y = 1 To NumOfTempRecs
      Get TempHandle, y, TempRec
        If CDbl(TempRec.CustomerNumber) = IdxRecs(x) Then
          Nextx = Nextx + 1
          ReDim Preserve PrintIdx(1 To Nextx) As Double
          PrintIdx(Nextx) = y 'Val(TempRec.CustomerNumber)
          Exit For
        End If
    Next y
      frmBLShowPctComp.ShowPctComp x, NumOfCustIdxRecs
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdAlign.Enabled = True
  cmdHelp.Enabled = True
    
  KillFile "artmplic.dat"
  OpenTempLicPrint TPHandle 'data reprints come from this
  TempPrintNum = 1
  
  frmBLShowPctComp.Label1 = "Printing Customer Business Licenses"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdAlign.Enabled = False
  cmdHelp.Enabled = False
  frmBLShowPctComp.cmdCancel.Visible = False
  
  If InStr(fpcmbBalanceType.Text, "Only") Then
    BalanceFlag = 1
  Else
    BalanceFlag = 2
  End If
  
  If PrintFees = False Then BalanceFlag = 1
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, PrintIdx(x), TempRec
    IssFee = TownRec.IssFee
    Get CustHandle, Val(TempRec.CustomerNumber), CustRec
      If UsePermLicNum = True Then LICENSE# = Val(CustRec.LICENSE)
      CustomerNumber = Val(TempRec.CustomerNumber) 'IdxRecs(x)
      For ll = 1 To 5
        Print #RptHandle,
      Next ll
      DidCnt = DidCnt + 1
      Print #RptHandle, Tab(11); CStr(SeqNum); Tab(37 - tab1); SHeading1$
      Print #RptHandle, Tab(37 - tab2); SHeading2$
      Print #RptHandle, Tab(37 - Tab3); SHeading3$
      Print #RptHandle, Tab(37 - Tab4); SHeading4$
      Print #RptHandle, Tab(66); fpBLYear.Text
      If CustRec.Prorate < 100 Then
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber)); Tab(26); "Fee prorated at " + CStr(CustRec.Prorate) + "%"
      Else
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber))
      End If
      Print #RptHandle, Tab(11); QPTrim$(CustRec.BillName)
      Print #RptHandle, Tab(11); QPTrim$(CustRec.ADDRESS1); Tab(58); Using("#######0", LICENSE#)
      Print #RptHandle, Tab(11); CustRec.ADDRESS2
      Print #RptHandle, Tab(11); RTrim$(CustRec.City); "  "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
      Print #RptHandle, Tab(55); FromDate$;
      Print #RptHandle, Tab(64); ExpireDate$
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, Tab(11); QPTrim$(CustRec.CustName)
      Print #RptHandle,
      Print #RptHandle,
      SCnt = 23
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT1)) = 0 Then GoTo To2
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT1);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1);
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.CatFee1)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1)
      End If
      SCnt = SCnt + 1
To2:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo To3 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT2);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2);
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.CatFee2)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2)
      End If
      SCnt = SCnt + 1
To3:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo To4 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT3);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3);
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.CatFee3)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3)
      End If
      SCnt = SCnt + 1
To4:
     If GetCatRecNum(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo To5 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT4);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4);
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.CatFee4)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4)
      End If
      SCnt = SCnt + 1
To5:
     If GetCatRecNum(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT5);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5);
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.CatFee5)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5)
      End If
      SCnt = SCnt + 1

ExitFormPrint1:
      If IssFee > 0 And PrintFees = True Then
        Print #RptHandle, Tab(15); "ISSUE FEE"; Tab(62); Using("####0.00", OldRound(IssFee))
        SCnt = SCnt + 1
      End If
      For LCnt = SCnt To 31
        Print #RptHandle,
      Next
      Print #RptHandle, ""

      For LCnt = 33 To 35
        Print #RptHandle, ""
      Next LCnt
      'Calc Total License Amount Here
      TotalBillAmt# = OldRound(TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5)
      TotalBillAmt# = OldRound(TotalBillAmt# + IssFee)
      If PrintFees = True Then
        Print #RptHandle, Tab(62); Using("####0.00", TotalBillAmt#) ' - OldRound(CustRec.AcctBal))
      Else
        Print #RptHandle, "No"
      End If
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      If BalanceFlag = 1 Then
        Print #RptHandle, Tab(62); Using("####0.00", TotalBillAmt#)
      Else
        Print #RptHandle, Tab(62); Using("####0.00", TempRec.AcctBal)
      End If
      Print #RptHandle,
      Print #RptHandle, "~"
      GoSub Post2TempAccount
      If UsePermLicNum = False Then LICENSE# = LICENSE# + 1
      frmBLShowPctComp.ShowPctComp x, NumOfTempRecs
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdAlign.Enabled = True
  cmdHelp.Enabled = True
  
  Print #RptHandle, Chr$(12);
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now

  ViewPrint ReportFile$, "Business License Printing", True
  KillFile ReportFile$
  
  ThisDate = fptxtVThru.Text
  ThisLen = Len(ThisDate)
  DHandle = FreeFile
  Open "validthrudate.dat" For Output As DHandle Len = ThisLen
  Print #DHandle, ThisDate
  Close DHandle
  
  ThisHeader = fptxtHeading(0).Text
  ThisLen = Len(ThisHeader)
  DHandle = FreeFile
  Open "appheader.dat" For Output As DHandle Len = ThisLen
  Print #DHandle, ThisHeader
  Close DHandle
  MainLog ("Business license tractor fed forms printed.")
  
  Exit Sub

Post2TempAccount: 'reprints capability depends on this code
  TempLPRec.LicNum = LICENSE# 'save for reprint purposes
  TempLPRec.RecNum = PrintIdx(x) 'Val(TempRec.CustomerNumber)
  TempLPRec.Expire = ExpireDate$
  TempLPRec.Head1 = QPTrim$(fptxtHeading(0))
  TempLPRec.Head2 = QPTrim$(fptxtHeading(1))
  TempLPRec.Head3 = QPTrim$(fptxtHeading(2))
  TempLPRec.Head4 = QPTrim$(fptxtHeading(3))
  TempLPRec.Issue = FromDate$
  TempLPRec.ThisYear = fpBLYear.Text$
  TempLPRec.SeqNum = SeqNum
  TempLPRec.FeeYN = False
  If PrintFees = True Then
    TempLPRec.FeeYN = True
  End If
  If BalanceFlag = 2 Then
    TempLPRec.TBalYN = True
  Else
    TempLPRec.TBalYN = False
  End If
  Put TPHandle, TempPrintNum, TempLPRec
  TempPrintNum = TempPrintNum + 1
  SeqNum = SeqNum + 1
  ' Update Customer Information First
  TempRec.LICENSE = LTrim$(Str$(LICENSE#))
  TempRec.VALID = Date2Num%(StoreExpireDate$)
  ' Calc New Running Balance
  ' Post Transaction Record First
  TempRec.TransDate = Date2Num%(PostDate$)
  If CustRec.FirstTrans = 0 Then
    TempRec.Prev = 0
  Else
    TempRec.Prev = CustRec.LastTrans
  End If
  
  Put TempHandle, PrintIdx(x), TempRec
  
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLic", "PrintText", Erl)
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

Private Sub LoadMe()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim ThisZip$
  Dim ThisYear As Integer
  Dim NextDate As Integer
  Dim NewYear$
  Dim DHandle As Integer
  Dim ThisDate$
  Dim ThisHeader$
  
  lblBalloon.Visible = False
'  fptxtVThru.ToolTipText = "The date entered here will appear on the license as the expiration date."
'  fptxtIssueDate.ToolTipText = "Enter the date which will appear on the business licenses the first valid day of this license period."
'  fptxtBegNum.ToolTipText = "Enter the new business license number that will begin the license printing process."
'  cmdList.ToolTipText = "Use this button to bring up a list of all customer's and their license numbers."
'  fptxtHeading(0).ToolTipText = "Optional line of text that will appear as the first line of the license header."
'  fptxtHeading(1).ToolTipText = "Optional line of text that will appear as the second line of the license header."
'  fptxtHeading(2).ToolTipText = "Optional line of text that will appear as the third line of the license header."
'  fptxtHeading(3).ToolTipText = "Optional line of text that will appear as the fourth line of the license header."
'  fpcmbFeeYN.ToolTipText = "Select Yes and any fees calculated and not posted for these licenses will be reset to zero. Choose No to allow the unposted calculations to remain."
'  fpcmbPrintFeesYN.ToolTipText = "This option allows the current fees to appear on each license. This option is disabled if the 'Charge Account With Fee Y/N' option is No."
'  fpcmbBalanceType.ToolTipText = "Business licenses can be printed with total outstanding balances and current balances or just current balances."
'  cmdAlign.ToolTipText = "Use this button to help line up license forms."
'  cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'  cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  UsePermLicNum = False
  
  If Exist("validthrudate.dat") Then
    DHandle = FreeFile
    Open "validthrudate.dat" For Input As #DHandle
    Line Input #DHandle, ThisDate
    fptxtVThru = ThisDate
    Close DHandle
  Else
    fptxtVThru = Date
    NewYear = fptxtVThru.AdjustDate(fptxtVThru.DateValue, 1, 0, 0)
    fptxtVThru.DateValue = NewYear
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  If Exist("appheader.dat") Then
    DHandle = FreeFile
    Open "appheader.dat" For Input As #DHandle
    Line Input #DHandle, ThisHeader
    fptxtHeading(0).Text = ThisHeader
    Close DHandle
  Else
    fptxtHeading(0).Text = QPTrim$(TownRec.TownName)
  End If
  
  If TownRec.LicNumPermYN = "Yes" Then
    UsePermLicNum = True
    fptxtBegNum.Enabled = False
    cmdList.Enabled = False
    fptxtBegNum.Text = "PERMANENT"
  Else
    fptxtBegNum.Text = FirstLicenseNum + 1
  End If
  
  fptxtFromDate = Date
'  fptxtVThru = Date
'  NewYear = fptxtVThru.AdjustDate(fptxtVThru.DateValue, 1, 0, 0)
'  fptxtVThru.DateValue = NewYear
  fptxtIssDate = Date
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fpcmbPrintFeesYN.Text = "Yes"
  fpcmbPrintFeesYN.AddItem "No"
  fpcmbPrintFeesYN.AddItem "Yes"
  fpcmbBalanceType.Text = "Current Balance Only"
  fpcmbBalanceType.AddItem "Current Balance Only"
  fpcmbBalanceType.AddItem "Total Balance"
  fptxtHeading(1).Text = QPTrim$(TownRec.TownAdd1)
  fptxtHeading(2).Text = QPTrim$(TownRec.TownAdd2)
  If Mid(TownRec.ZipCode, 7, 1) = " " Then
    ThisZip = Mid(TownRec.ZipCode, 1, 5)
  Else
    ThisZip = QPTrim$(TownRec.ZipCode)
  End If
  fptxtHeading(3).Text = QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + "  " + ThisZip
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtHeading_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 3 Then
    If KeyCode = vbKeyDown Then
      fptxtIssDate.SetFocus
    End If
  End If
End Sub

Private Function Look4DupLicNums() As Boolean
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim x As Double, y As Double
  Dim ThisLic As Double
  Dim ThatLic As Double
  Dim YCnt As Double
  
  'this function takes the beginning license number entered
  'in the 'Beginning License Number' field and checks the rest
  'of the customers who are not included in this license
  'processing to make sure a duplicate license number will
  'not be assigned
  On Error GoTo ERRORSTUFF
  Look4DupLicNums = False
  
  ThisLic = CDbl(fptxtBegNum.Text)
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
      If QPTrim$(CustRec.IssueLicense) = "Y" Then
        YCnt = YCnt + 1
      End If
  Next x
  
  If YCnt = 0 Then
    Close CustHandle
    Exit Function
  End If
  
  ReDim YCntIdx(1 To YCnt) As String
  YCntIdx(1) = ThisLic
  For x = 2 To YCnt
    ThisLic = ThisLic + 1
    YCntIdx(x) = ThisLic
  Next x
  
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
      If QPTrim(CustRec.LICENSE) = "" Then GoTo NoLicenseNum
      If QPTrim(CustRec.Deleted) <> "Y" And QPTrim$(CustRec.IssueLicense) = "N" Then
        ThatLic = CDbl(CustRec.LICENSE)
        For y = 1 To YCnt
          If YCntIdx(y) = ThatLic Then
            frmBLMessageBoxJr.Label1.Caption = "The beginning license number entered would cause a duplicate license number problem between new License # " + CStr(YCntIdx(y)) + " and the existing license number of current customer " + QPTrim(CustRec.CustName) + " who is not included in this license process. Please revise your beginning license number to avoid this conflict."
            frmBLMessageBoxJr.Label1.Top = 430
            frmBLMessageBoxJr.Label1.Height = 1300
            frmBLMessageBoxJr.Show vbModal
            fptxtBegNum.SetFocus
            Look4DupLicNums = True
            GoTo DoneHere
          End If
        Next y
      End If
NoLicenseNum:
  Next x
  
DoneHere:

  Exit Function
  
ERRORSTUFF:
  frmBLMessageBoxJr.Label1.Caption = "ERROR: An error has occurred in the 'Look4DupLicNum' function for customer number " + QPTrim$(CustRec.CustNumb) + "."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  Close CustHandle
  
End Function

Private Sub fpcmbPrintFeesYN_Change()
  If QPTrim$(fpcmbPrintFeesYN.Text) = "" Then
    fpcmbPrintFeesYN.Text = "Yes"
  End If
  If QPTrim$(fpcmbPrintFeesYN.Text) = "No" Then
    fpcmbBalanceType.Enabled = False
  Else
    fpcmbBalanceType.Enabled = True
  End If
End Sub

