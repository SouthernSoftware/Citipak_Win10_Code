VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmESCQrtRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ESC Quarterly Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmESCQrtRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6015
      Left            =   2100
      TabIndex        =   6
      Top             =   1440
      Width           =   7455
      _Version        =   196609
      _ExtentX        =   13150
      _ExtentY        =   10610
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmESCQrtRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboDiskFile 
         Height          =   405
         Left            =   5565
         TabIndex        =   4
         Top             =   3600
         Width           =   780
         _Version        =   196608
         _ExtentX        =   1376
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmESCQrtRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPayType 
         Height          =   405
         Left            =   2595
         TabIndex        =   3
         Top             =   3045
         Width           =   4275
         _Version        =   196608
         _ExtentX        =   7541
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmESCQrtRpt.frx":0BDD
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3480
         TabIndex        =   5
         Top             =   4440
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmESCQrtRpt.frx":0ED4
      End
      Begin VB.CheckBox chkSaveToFolder 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Save File to Citipak Folder Instead of Drive A"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   4080
         Visible         =   0   'False
         Width           =   3735
      End
      Begin EditLib.fpDateTime fptxtYear 
         Height          =   390
         Left            =   3555
         TabIndex        =   1
         Top             =   1965
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2138
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
         UserEntry       =   1
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
         Text            =   "2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "YYYY"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtQtr 
         Height          =   390
         Left            =   4755
         TabIndex        =   0
         Top             =   1410
         Width           =   630
         _Version        =   196608
         _ExtentX        =   1122
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4176
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the quarterly ESC report."
         Top             =   5040
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmESCQrtRpt.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1248
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   5040
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmESCQrtRpt.frx":13AA
      End
      Begin EditLib.fpText fptxtMaxSal 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   390
         Left            =   3405
         TabIndex        =   2
         Top             =   2520
         Width           =   2190
         _Version        =   196608
         _ExtentX        =   3863
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 , $ ."
         MaxLength       =   11
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Max Salary:"
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
         Left            =   1725
         TabIndex        =   15
         Top             =   2640
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
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
         Left            =   1605
         TabIndex        =   12
         Top             =   4530
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Create a magnetic disk while processing?"
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
         Left            =   960
         TabIndex        =   11
         Top             =   3696
         Width           =   4572
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Wage Preference:"
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
         Left            =   435
         TabIndex        =   10
         Top             =   3150
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "ESC Quarterly Wage Report"
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
         Height          =   444
         Left            =   1728
         TabIndex        =   9
         Top             =   576
         Width           =   4044
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Using Quarter:"
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
         Left            =   2010
         TabIndex        =   8
         Top             =   1560
         Width           =   2550
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Year:"
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
         Left            =   2355
         TabIndex        =   7
         Top             =   2070
         Width           =   780
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   1536
         Top             =   384
         Width           =   4428
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6264
      Left            =   1980
      Top             =   1302
      Width           =   7692
   End
End
Attribute VB_Name = "frmESCQrtRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Dim RemitNumb$
  Dim EmplrAcct$
  Dim Gross$
  Dim SOCGrossFlag As Boolean
  Dim PrintLine() As ESCMAG2DiskType
  Dim PrintLineCnt As Integer
  Private Temp_Class As Resize_Class

Private Sub cmdProcess_Click()
  Call SaveESCMaxPay
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadESCScreen
  Me.HelpContextID = hlpESCReport
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmESCQrtRpt
End Sub

Private Sub LoadESCScreen()
   Dim Today As String * 11
   Dim x As Integer
   Dim MaxSal As Double
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtYear.Text = Mid(Today, 7, 4)
   fpcomboPayType.Text = "Gross Wage"
   fpcomboPayType.AddItem "Gross Wage w/o SS Exempt Deductions"
   fpcomboPayType.AddItem "Social Security Wage"
   fpcomboPayType.AddItem "Retirement Wage"
   fpcomboPayType.AddItem "Gross Wage"
   fpcomboDiskFile.Text = "N"
   fpcomboDiskFile.AddItem "Y"
   fpcomboDiskFile.AddItem "N"
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
   MaxSal = 0
   MaxSal = LoadESCMaxPay
   fptxtMaxSal.Text = FormatCurrency(MaxSal, 2, vbTrue)
End Sub

Private Sub PrintGraphics()

  Dim RptQuarter$, Qtr$, RptTitle$, RptName$
  Dim Year As String
  Dim CrLf2$, CrLf5$, CrLf6$, fmt$, Fmt2$, CrLf$, CrLf8$
  Dim MaxLines As Integer, IdxRecLen As Integer
  Dim UnitHandle As Integer, IdxFileSize&
  Dim EmpRecSize As Long, TRecSize As Long
  Dim NumOfRecs As Long, cnt As Integer
  Dim RHandle As Integer, THandle As Integer, DHandle As Integer
  Dim LineCnt As Integer
  Dim RecNo As Long, TransRecNum&, GrandTotal#
  Dim DoQtrLine As Boolean, FF$, PageNo As Integer
  Dim GrossOvr#, TotalGrossOvr#, Cnt2 As Integer
  Dim YTD2PrevQtr#, YTD2ThisQtr#, SubTotal#
  Dim EmpIdxLNameHandle As Integer, x As Integer
  Dim NameIdxLName As NameSortIdxType
  Dim Emp2Rec As EmpData2Type
  Dim dlm$
  Dim DedRec As DedCodeRecType
  Dim NumOfDedRecs As Integer
  Dim DedHandle As Integer
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  If fptxtYear.Text = "" Then
     MsgBox "Please enter a Year"
     fptxtYear.SetFocus
     Exit Sub
  End If

  If Val(fptxtYear.Text) < 1920 Or Val(fptxtYear.Text) > 2099 Then
     MsgBox "Please enter a valid Year (####)"
     fptxtYear.SetFocus
     Exit Sub
  End If

  'the next if should never happen because the allowable
  'values in fptxtQtr are 1 2 3 & 4 only
  If Val(fptxtQtr.Text) < 1 Or Val(fptxtQtr.Text) > 4 Then
     MsgBox "Please enter a valid Quarter value"
     fptxtQtr.SetFocus
     Exit Sub
  End If
  
  RptQuarter$ = QPTrim$(fptxtQtr.Text)
  GlblQtr$ = RptQuarter$ 'GlblQtr passes the quarter to the ar report
  Year$ = QPTrim$(fptxtYear.Text)

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim GrsRpt(1 To 3) As ESCGrossWageRptType
  ReDim Line2(1) As String * 80

  ReDim EQtrWage(1 To 4) As Double
  ReDim QtrDates(1 To 4) As QtrDateType

  ReDim ToDisk1(1) As ESC2DiskRecType1
  ToDisk1(1).Seasonal = ""
  ToDisk1(1).Fill1 = ""
  ToDisk1(1).CrLf = CrLf$

  fmt$ = "#,###,##0.00"
  Fmt2$ = "###,###,##0.00"

  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
    
  QtrDates(1).LDate = Date2Num("01/01/" + Year$)
  QtrDates(1).HDate = Date2Num("03/31/" + Year$)
  QtrDates(2).LDate = Date2Num("04/01/" + Year$)
  QtrDates(2).HDate = Date2Num("06/30/" + Year$)
  QtrDates(3).LDate = Date2Num("07/01/" + Year$)
  QtrDates(3).HDate = Date2Num("09/30/" + Year$)
  QtrDates(4).LDate = Date2Num("10/01/" + Year$)
  QtrDates(4).HDate = Date2Num("12/31/" + Year$)
  Qtr$ = QPTrim$(RptQuarter) + " " + Year$
  RptTitle$ = "ESC Quarterly Wage Report"
  
  If fpcomboDiskFile.Text = "Y" Then
    Call ESC2Disk
  End If

  RptName$ = "PRRPTS\ESCQTR" + QPTrim$(RptQuarter) + ".RPT"
  On Error GoTo ErrorHandler
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) \ 2
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "ESC Quarterly Wage Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If (Emp2Rec.LastTransRec <= 0) Or (Emp2Rec.ExcludeESC = "Y") Then
      GoTo SkipEm5
    End If
    TransRecNum& = Emp2Rec.LastTransRec
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      GoSub SumEmpESCData
      If TransHRec(1).PrevTransRec <= 0 Then
        GoSub PrintEmpESCLine
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm5:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
Next RecNo

  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  '               0        1
  Print #RHandle, ""; dlm; ""; dlm;
  '                2        3
  Print #RHandle, ""; dlm; ""; dlm;
  '                4       5        6
  Print #RHandle, ""; dlm; ""; dlm; ""; dlm;
  '
  If Unit(1).ESCRTYPE = 2 Then
  '                              7                                   8                                      9
    Print #RHandle, Using$(Fmt2$, GrandTotal#); dlm; Using$(Fmt2$, TotalGrossOvr#); dlm; Using$(Fmt2$, OldRound(GrandTotal# - TotalGrossOvr#))
  Else
    Print #RHandle, Using$(Fmt2$, GrandTotal#); dlm; ""; dlm; ""
  End If
  Close DHandle
  Close THandle
  Close RHandle
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arESCRpt.Show
  frmLoadingRpt.Show
  MainLog ("ESC Quarterly Wage Report processed.")
  
Exit Sub

SumEmpESCData:
  For Cnt2 = 1 To 4  'put gross into correct quarter
    If (TransHRec(1).CheckDate >= QtrDates(Cnt2).LDate) And (TransHRec(1).CheckDate <= QtrDates(Cnt2).HDate) Then
      Select Case fpcomboPayType.Text
      Case "Gross Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case "Gross Wage w/o SS Exempt Deductions"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        OpenDedCodeFile DedHandle
        NumOfDedRecs = LOF(DedHandle) / Len(DedRec)
        For x = 1 To NumOfDedRecs
          Get DedHandle, x, DedRec
          If QPTrim$(DedRec.DCSOC1) = "Y" Then
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) - TransHRec(1).DAmt(x))
          End If
        Next x
        Close DedHandle
        Exit For
      Case "Social Security Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).SocGrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case "Retirement Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).RetGrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case Else
        MsgBox "Please make a selection in the Wage Preference list box."
        fpcomboPayType.SetFocus
        Exit Sub
      End Select
    End If
  Next
Return

PrintEmpESCLine:

  If Unit(1).ESCRTYPE = 2 Then
    If RptQuarter > 1 Then            'if Not 1st qtr, we will have to
      For Cnt2 = 1 To RptQuarter - 1  'examine gross thru prior qtr
        YTD2PrevQtr# = OldRound(YTD2PrevQtr# + EQtrWage(Cnt2))
      Next
      For Cnt2 = 1 To RptQuarter
        YTD2ThisQtr# = OldRound(YTD2ThisQtr# + EQtrWage(Cnt2))
      Next
      If YTD2PrevQtr# > Unit(1).TAXWBASE Then     'if the prior qtr was
        GrossOvr# = EQtrWage(RptQuarter)          'over also TaxBase also
      ElseIf YTD2ThisQtr# > Unit(1).TAXWBASE Then         'else if gross thru
        GrossOvr# = OldRound(YTD2ThisQtr# - Unit(1).TAXWBASE) 'this qtr is over
      Else   'still not over
        GrossOvr# = 0
      End If
'*** This deals with the 1ST quarter only!!
    ElseIf EQtrWage(1) > Unit(1).TAXWBASE Then  'else this is 1st qtr report
      GrossOvr# = OldRound(EQtrWage(1) - Unit(1).TAXWBASE)
    Else
      GrossOvr# = 0
    End If
'*** 1ST Quarter end
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter) 'OldRound(EQtrWage(RptQuarter) - GrossOvr#)
  Else       'not type 2 report
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter)
  End If

  TotalGrossOvr# = OldRound(TotalGrossOvr# + GrossOvr#)
  GrandTotal# = OldRound(GrandTotal# + GrsRpt(1).GrossPay)

  If GrsRpt(1).GrossPay = 0 Then GoTo SkipEMPPrint

  RSet ToDisk1(1).GPay = Using$(fmt$, GrsRpt(1).GrossPay)
  LSet ToDisk1(1).ESSN = Left$(Emp2Rec.EmpSSN, 3) + "-" + Mid$(Emp2Rec.EmpSSN, 4, 2) + "-" + Mid$(Emp2Rec.EmpSSN, 6, 4)
  LSet ToDisk1(1).EName = Left$(Emp2Rec.EmpFName, 1) + "  " + Left$(Emp2Rec.EmpLName, 18)

  If DoQtrLine Then
    RSet ToDisk1(1).Qtr = Qtr$
    DoQtrLine = False
  Else
    RSet ToDisk1(1).Qtr = " "
  End If
  ThisCnt = ThisCnt + 1
  '                           0                           1
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; QPTrim$(Unit(1).ESCEmplrNum); dlm; 'Unit(1).ESCRemitNum...Unit(1).ESCEmplrNum
  '                          2                3               4
  Print #RHandle, QPTrim$(RptQuarter); dlm; Year$; dlm; ToDisk1(1).ESSN; dlm;
  '                         5                   6                          7
  Print #RHandle, ToDisk1(1).EName; dlm; ToDisk1(1).GPay; dlm;
  '
  If Unit(1).ESCRTYPE = 2 Then
  '                               7                           8                                       9
    Print #RHandle, Using$(Fmt2$, GrandTotal#); dlm; Using$(Fmt2$, TotalGrossOvr#); dlm; Using$(Fmt2$, OldRound(GrandTotal# - TotalGrossOvr#))
  Else
  '                               7                           8                                       9
    Print #RHandle, Using$(Fmt2$, GrandTotal#); dlm; ""; dlm; ""
  End If
  
SkipEMPPrint:
  GrsRpt(1) = GrsRpt(2)
  YTD2PrevQtr# = 0
  YTD2ThisQtr# = 0
  For Cnt2 = 1 To 4
    EQtrWage(Cnt2) = 0
  Next

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."


DedExitRpt:

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmESCQrtRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub ESC2Disk()
  Dim ESCRecLen As Long
  Dim TRecLen As Long
  Dim EmpRecLen As Long
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Long
  Dim EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer
  Dim ESCReport$
  Dim RptFile As Integer
  Dim ESCFile As Integer
  Dim TRFile As Integer
  Dim EmpFile As Integer
  Dim RecNo As Long, Quarter$
  Dim EscExcl As Long
  Dim EmpCnt As Long
  Dim CPos As Integer, Cnt2 As Integer
  Dim LName$, TransRecNum&
  Dim Year$, x As Integer
  Dim ESCHandle As Integer
  Dim DedRec As DedCodeRecType
  Dim NumOfDedRecs As Integer
  Dim DedHandle As Integer
  Dim ThisCnt As Integer
  Dim YrEnd As Integer
  Dim YrStart As Integer
  Dim Month1EmpPayTotal As Double
  Dim Month2EmpPayTotal As Double
  Dim Month3EmpPayTotal As Double
  Dim Qtr1TotalEmpPay As Double
  Dim Qtr2TotalEmpPay As Double
  Dim Qtr3TotalEmpPay As Double
  Dim Qtr4TotalEmpPay As Double
  Dim Qtr1TotalPay As Double
  Dim Qtr2TotalPay As Double
  Dim Qtr3TotalPay As Double
  Dim Qtr4TotalPay As Double
  Dim Qtr1ExcessPay As Double
  Dim Qtr2ExcessPay As Double
  Dim Qtr3ExcessPay As Double
  Dim Qtr4ExcessPay As Double
  Dim Qtr1TaxablePay As Double
  Dim Qtr2TaxablePay As Double
  Dim Qtr3TaxablePay As Double
  Dim Qtr4TaxablePay As Double
  Dim HoldAmt As Double
  Dim FMonth12 As Integer
  Dim SMonth12 As Integer
  Dim TMonth12 As Integer
  Dim MaxPay As Double
  Dim Month1EmpCnt As Integer
  Dim Month2EmpCnt As Integer
  Dim Month3EmpCnt As Integer
  Dim WhichQtr As Integer
  Dim Month1Employment As String * 5
  Dim Month2Employment As String * 5
  Dim Month3Employment As String * 5
  Dim QtrTotalWages As String * 11
  Dim QtrExcessWages As String * 11
  Dim QtrTaxableWages As String * 11
  Dim ReportingQtr As String * 1
  Dim Blank12 As String * 12
  Dim Blank80 As String * 80
  Dim Header As String
  Dim LF As String * 2
  Dim Qtr1ExcessByEmp As Double
  Dim Qtr2ExcessByEmp As Double
  Dim Qtr3ExcessByEmp As Double
  Dim Qtr4ExcessByEmp As Double
  Dim Qtr1TaxableByEmp As Double
  Dim Qtr2TaxableByEmp As Double
  Dim Qtr3TaxableByEmp As Double
  Dim Qtr4TaxableByEmp As Double
  Dim NoADrive As Boolean
  Dim FileDir As String
'  Dim AHandle As Integer
'
'  AHandle = FreeFile
'  Open "ESCTest.txt" For Output As AHandle
  NoADrive = False
  PrintLineCnt = 0
  Blank80 = Str(80)
  MaxPay = CDbl(fptxtMaxSal.Text)
  Year$ = QPTrim$(fptxtYear.Text)
  Select Case CInt(fptxtQtr.Text)
    Case 1
      FMonth12 = Date2Num("01/12/" + Year$)
      SMonth12 = Date2Num("02/12/" + Year$)
      TMonth12 = Date2Num("03/12/" + Year$)
    Case 2
      FMonth12 = Date2Num("04/12/" + Year$)
      SMonth12 = Date2Num("05/12/" + Year$)
      TMonth12 = Date2Num("06/12/" + Year$)
    Case 3
      FMonth12 = Date2Num("07/12/" + Year$)
      SMonth12 = Date2Num("08/12/" + Year$)
      TMonth12 = Date2Num("09/12/" + Year$)
    Case 4
      FMonth12 = Date2Num("10/12/" + Year$)
      SMonth12 = Date2Num("11/12/" + Year$)
      TMonth12 = Date2Num("12/12/" + Year$)
  End Select
  
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim ESCRec(1) As ESCMAG2DiskType
  ReDim QtrDates(1 To 4) As QtrDateType
  ReDim EmpNum1(1 To 1) As String
  Dim EmpNum1Cnt As Integer
  ReDim EmpNum2(1 To 1) As String
  Dim EmpNum2Cnt As Integer
  ReDim EmpNum3(1 To 1) As String
  Dim EmpNum3Cnt As Integer
  ReDim PrintLine(1 To 1) As ESCMAG2DiskType
  
  QtrDates(1).LDate = Date2Num("01/01/" + Year$)
  QtrDates(1).HDate = Date2Num("03/31/" + Year$)
  QtrDates(2).LDate = Date2Num("04/01/" + Year$)
  QtrDates(2).HDate = Date2Num("06/30/" + Year$)
  QtrDates(3).LDate = Date2Num("07/01/" + Year$)
  QtrDates(3).HDate = Date2Num("09/30/" + Year$)
  QtrDates(4).LDate = Date2Num("10/01/" + Year$)
  QtrDates(4).HDate = Date2Num("12/31/" + Year$)
  YrStart = Date2Num("01/01/" + Year$)
  YrEnd = Date2Num("12/31/" + Year$)
  ESCRecLen = Len(ESCRec(1))
  TRecLen = Len(TransHRec(1))
  EmpRecLen = Len(Emp2Rec(1))

  EmpIdxLNameHandle = FreeFile
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle

  RemitNumb$ = QPTrim$(Unit(1).ESCRemitNum)
  EmplrAcct$ = QPTrim$(Unit(1).ESCEmplrNum)
  
  If Len(RemitNumb$) = 0 Or QPTrim$(RemitNumb$) = "0" Then
    MsgBox "Please enter a ESC Remit Number in the Employer File located on the Control Maintenance Menu."
    Unload frmProcessing
    MsgBox "No data saved." ' to drive A:"
    Close
    Exit Sub
  End If
  
  If Len(EmplrAcct$) = 0 Or QPTrim$(EmplrAcct$) = "0" Then
    MsgBox "Please enter a ESC Employer Number in the Employer File located on the Control Maintenance Menu."
    Unload frmProcessing
    MsgBox "No data saved." ' to drive A:"
    Close
    Exit Sub
  End If
  
'  If Len(RemitNumb$) <> 6 Then
'    MsgBox "Please enter a 6 character ESC Remit Number in the Employer File located on the Control Maintenance Menu."
'    Unload frmProcessing
'    MsgBox "No data saved." ' to drive A:"
'    Close
'    Exit Sub
'  End If
'
'  If Len(EmplrAcct$) <> 7 Then
'    MsgBox "Please enter a 7 character ESC Employer Number in the Employer File located on the Control Maintenance Menu."
'    Unload frmProcessing
'    MsgBox "No data saved." ' to drive A:"
'    Close
'    Exit Sub
'  End If
  
  On Local Error GoTo ErrorHandler

  If chkSaveToFolder.Value = 1 Then
    FileDir = App.Path + "\NCESC"
    If Not DirExists(FileDir) Then
      frmMessageWOpts.Label1.Caption = "The directory 'NCESC' could not be located in the Citipak directory. OK to create it now?"
      frmMessageWOpts.Label1.Top = 800
      frmMessageWOpts.cmdCont.Text = "F10 Make NCESC"
      frmMessageWOpts.cmdExit.Text = "ESC Escape"
      frmMessageWOpts.Show vbModal
      If frmMessageWOpts.fptxtChoice.Text = "continue" Then
        Unload frmMessageWOpts
        MkDir App.Path + "\NCESC"
      Else
        Unload frmMessageWOpts
        MsgBox ("Save to Citipak\NCESC\ESCNC.WGS aborted.")
        Close
        Exit Sub
      End If
    End If
    'MsgBox "Saving to Citipak\ESCNC\ESCNC.WGS", vbInformation, "Citipak"
    DoEvents
    frmProcessing.Label1.Caption = "Saving" ' to Citipak\ESCNC\ESCNC.WGS"
    DoEvents
    frmProcessing.Show , Me
    DoEvents
    ESCReport$ = FileDir + "\ESCNC.WGS"
    RptFile = FreeFile
    Open ESCReport$ For Output As RptFile
    Close RptFile
    ESCReport$ = FileDir + "\ESCNC.WGS"
  Else
    NoADrive = True
    frmProcessing.Label1.Caption = "Saving to Drive A:"
    DoEvents
    frmProcessing.Show , Me
    DoEvents
    ESCReport$ = "A:\ESCNC.WGS"
    RptFile = FreeFile
    Open ESCReport$ For Output As RptFile
    Close RptFile
    NoADrive = False
  End If
  
  ESCFile = FreeFile
  Open ESCReport$ For Random As #ESCFile Len = ESCRecLen
  OpenTransHistFile TRFile
  OpenEmpData2File EmpFile
  If QPTrim$(fpcomboPayType.Text) = "Gross Wage w/o SS Exempt Deductions" Then
    OpenDedCodeFile DedHandle
    NumOfDedRecs = LOF(DedHandle) / Len(DedRec)
  End If
  For RecNo = 1 To NumOfRecs
    Get #EmpFile, RecNo, Emp2Rec(1)
    'If QPTrim$(ReplaceString(Emp2Rec(1).EmpSSN, "-", "")) = "001529572" Then Stop
    If Emp2Rec(1).ExcludeESC = "Y" Then
      EscExcl = EscExcl + 1
      GoTo SkipEm
    End If
    
    If Emp2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEm
    End If

    ReDim EQtrWage(1 To 4) As Double
    TransRecNum& = Emp2Rec(1).LastTransRec
    
    Do
      Get #TRFile, TransRecNum&, TransHRec(1)
      GoSub SumESCData
      If (TransHRec(1).CheckDate >= QtrDates(CInt(fptxtQtr.Text)).LDate) And (TransHRec(1).CheckDate <= QtrDates(CInt(fptxtQtr.Text)).HDate) Then
        Dim WhichMonth As Integer
        WhichMonth = GetMonth(MakeRegDate(TransHRec(1).CheckDate))
        Select Case WhichMonth
          Case 1
            If EmpNum1Cnt = 0 Then
              ReDim Preserve EmpNum1(1 To 1) As String
              EmpNum1(1) = Emp2Rec(1).EmpNo
              EmpNum1Cnt = EmpNum1Cnt + 1
              If Emp2Rec(1).EMPHDATE <= FMonth12 Then
                Month1EmpCnt = Month1EmpCnt + 1
              End If
            Else
              For x = 1 To EmpNum1Cnt
                If Emp2Rec(1).EmpNo = EmpNum1(x) Then
                  Exit For
                End If
              Next x
              If x > EmpNum1Cnt Then
                EmpNum1Cnt = EmpNum1Cnt + 1
                ReDim Preserve EmpNum1(1 To EmpNum1Cnt) As String
                EmpNum1(EmpNum1Cnt) = Emp2Rec(1).EmpNo
                If Emp2Rec(1).EMPHDATE <= FMonth12 Then
                  Month1EmpCnt = Month1EmpCnt + 1
                End If
              End If
            End If
            
          Case 2
            If EmpNum2Cnt = 0 Then
              ReDim Preserve EmpNum2(1 To 1) As String
              EmpNum2(1) = Emp2Rec(1).EmpNo
              EmpNum2Cnt = EmpNum2Cnt + 1
              If Emp2Rec(1).EMPHDATE <= SMonth12 Then
                Month2EmpCnt = Month2EmpCnt + 1
              End If
            Else
              For x = 1 To EmpNum2Cnt
                If Emp2Rec(1).EmpNo = EmpNum2(x) Then
                  Exit For
                End If
              Next x

              If x > EmpNum2Cnt Then
                EmpNum2Cnt = EmpNum2Cnt + 1
                ReDim Preserve EmpNum2(1 To EmpNum2Cnt) As String
                EmpNum2(EmpNum2Cnt) = Emp2Rec(1).EmpNo
                If Emp2Rec(1).EMPHDATE <= SMonth12 Then
                  Month2EmpCnt = Month2EmpCnt + 1
                End If
              End If
            End If
            
          Case 3
            If EmpNum3Cnt = 0 Then
              ReDim Preserve EmpNum3(1 To 1) As String
              EmpNum3(1) = Emp2Rec(1).EmpNo
              EmpNum3Cnt = EmpNum3Cnt + 1
              If Emp2Rec(1).EMPHDATE <= TMonth12 Then
                Month3EmpCnt = Month3EmpCnt + 1
              End If
            Else
              For x = 1 To EmpNum3Cnt
                If Emp2Rec(1).EmpNo = EmpNum3(x) Then
                  Exit For
                End If
              Next x
              If x > EmpNum3Cnt Then
                EmpNum3Cnt = EmpNum3Cnt + 1
                ReDim Preserve EmpNum3(1 To EmpNum3Cnt) As String
                EmpNum3(EmpNum3Cnt) = Emp2Rec(1).EmpNo
                If Emp2Rec(1).EMPHDATE <= TMonth12 Then
                  Month3EmpCnt = Month3EmpCnt + 1
                End If
              End If
            End If
              
        End Select
      End If
      ReDim EQtrWage2(1 To 4) As Double
      GoSub SumAnnualESCData
      Select Case WhichQtr
        Case 1
          Qtr1TotalEmpPay = Qtr1TotalEmpPay + EQtrWage2(1)
          Qtr1TotalPay = Qtr1TotalPay + EQtrWage2(1)
        Case 2
          Qtr2TotalEmpPay = Qtr2TotalEmpPay + EQtrWage2(2)
          Qtr2TotalPay = Qtr2TotalPay + EQtrWage2(2)
        Case 3
          Qtr3TotalEmpPay = Qtr3TotalEmpPay + EQtrWage2(3)
          Qtr3TotalPay = Qtr3TotalPay + EQtrWage2(3)
        Case 4
          Qtr4TotalEmpPay = Qtr4TotalEmpPay + EQtrWage2(4)
          Qtr4TotalPay = Qtr4TotalPay + EQtrWage2(4)
     End Select
      
      If TransHRec(1).PrevTransRec <= 0 Then
        Emp2Rec(1).EmpPin = Emp2Rec(1).EmpPin
        GoSub PrintESCLine
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
    'collect data for header row
    Qtr1TotalPay = OldRound(Qtr1TotalPay)
    Qtr2TotalPay = OldRound(Qtr2TotalPay)
    Qtr3TotalPay = OldRound(Qtr3TotalPay)
    Qtr4TotalPay = OldRound(Qtr4TotalPay)
    Qtr1ExcessByEmp = 0
    Qtr2ExcessByEmp = 0
    Qtr3ExcessByEmp = 0
    Qtr4ExcessByEmp = 0
    Qtr1TaxableByEmp = 0
    Qtr2TaxableByEmp = 0
    Qtr3TaxableByEmp = 0
    Qtr4TaxableByEmp = 0
    
    Dim MaxTot As Double
    MaxTot = MaxPay
    For x = 1 To 4
      Select Case x
        Case 1
          If Qtr1TotalEmpPay > MaxPay Then
            Qtr1ExcessByEmp = Qtr1ExcessByEmp + Qtr1TotalEmpPay - MaxTot
            Qtr1TaxableByEmp = Qtr1TaxableByEmp + MaxTot
            MaxTot = 0
          ElseIf Qtr1TotalEmpPay <= MaxTot Then
            Qtr1TaxableByEmp = Qtr1TaxableByEmp + Qtr1TotalEmpPay
            MaxTot = MaxTot - Qtr1TotalEmpPay 'MaxTot = amt paid toward maxpay
          End If
        Case 2
          If Qtr2TotalEmpPay > MaxTot Then
            Qtr2ExcessByEmp = Qtr2ExcessByEmp + Qtr2TotalEmpPay - MaxTot
            Qtr2TaxableByEmp = Qtr2TaxableByEmp + Qtr2TotalEmpPay - (Qtr2TotalEmpPay - MaxTot)
            If Qtr2TaxableByEmp < 0 Then Qtr2TaxableByEmp = 0
            MaxTot = 0
          ElseIf Qtr2TotalEmpPay <= MaxTot Then
            Qtr2TaxableByEmp = Qtr2TaxableByEmp + Qtr2TotalEmpPay
            MaxTot = MaxTot - Qtr2TotalEmpPay
          End If
        Case 3
          If Qtr3TotalEmpPay > MaxTot Then
            Qtr3ExcessByEmp = Qtr3ExcessByEmp + Qtr3TotalEmpPay - MaxTot
            Qtr3TaxableByEmp = Qtr3TaxableByEmp + Qtr3TotalEmpPay - (Qtr3TotalEmpPay - MaxTot)
            If Qtr3TaxableByEmp < 0 Then Qtr3TaxableByEmp = 0
            MaxTot = 0
          ElseIf Qtr3TotalEmpPay <= MaxTot Then
            Qtr3TaxableByEmp = Qtr3TaxableByEmp + Qtr3TotalEmpPay
            MaxTot = MaxTot - Qtr3TotalEmpPay
          End If
       Case 4
          If Qtr4TotalEmpPay > MaxTot Then
            Qtr4ExcessByEmp = Qtr4ExcessByEmp + Qtr4TotalEmpPay - MaxTot
            Qtr4TaxableByEmp = Qtr4TaxableByEmp + Qtr4TotalEmpPay - (Qtr4TotalEmpPay - MaxTot)
            If Qtr4TaxableByEmp < 0 Then Qtr4TaxableByEmp = 0
            MaxTot = 0
          ElseIf Qtr4TotalEmpPay <= MaxTot Then
            Qtr4TaxableByEmp = Qtr4TaxableByEmp + Qtr4TotalEmpPay
            MaxTot = MaxTot - Qtr4TotalEmpPay
          End If
      End Select
    Next x
    
    Qtr1ExcessPay = Qtr1ExcessPay + Qtr1ExcessByEmp
    Qtr2ExcessPay = Qtr2ExcessPay + Qtr2ExcessByEmp
    Qtr3ExcessPay = Qtr3ExcessPay + Qtr3ExcessByEmp
    Qtr4ExcessPay = Qtr4ExcessPay + Qtr4ExcessByEmp
    
    Qtr1TaxablePay = Qtr1TaxablePay + Qtr1TaxableByEmp
    Qtr2TaxablePay = Qtr2TaxablePay + Qtr2TaxableByEmp
    Qtr3TaxablePay = Qtr3TaxablePay + Qtr3TaxableByEmp
    Qtr4TaxablePay = Qtr4TaxablePay + Qtr4TaxableByEmp
    
    Qtr1TotalEmpPay = 0
    Qtr2TotalEmpPay = 0
    Qtr3TotalEmpPay = 0
    Qtr4TotalEmpPay = 0
    HoldAmt = 0
SkipEm:
  Next
  
  ReportingQtr = fptxtQtr.Text
  Month1Employment = RightJustify(Month1EmpCnt, 5)
  Month2Employment = RightJustify(Month2EmpCnt, 5)
  Month3Employment = RightJustify(Month3EmpCnt, 5)
  Blank12 = "            "
  Select Case CInt(fptxtQtr.Text)
    Case 1
     QtrTotalWages = RightJustify(Qtr1TotalPay, 11)
     QtrExcessWages = RightJustify(Qtr1ExcessPay, 11)
     QtrTaxableWages = RightJustify(Qtr1TaxablePay, 11)
    Case 2
     QtrTotalWages = RightJustify(Qtr2TotalPay, 11)
     QtrExcessWages = RightJustify(Qtr2ExcessPay, 11)
     QtrTaxableWages = RightJustify(Qtr2TaxablePay, 11)
    Case 3
     QtrTotalWages = RightJustify(Qtr3TotalPay, 11)
     QtrExcessWages = RightJustify(Qtr3ExcessPay, 11)
     QtrTaxableWages = RightJustify(Qtr3TaxablePay, 11)
    Case 4
     QtrTotalWages = RightJustify(Qtr4TotalPay, 11)
     QtrExcessWages = RightJustify(Qtr4ExcessPay, 11)
     QtrTaxableWages = RightJustify(Qtr4TaxablePay, 11)
  End Select
 
'  Dim ESC As ESCLineN
'  Dim ESCNHandle As Integer
'  Dim NumOfESCRecs As Integer
'  EmplrAcct = ReplaceString(EmplrAcct, "-", "")
'  EmplrAcct = ReplaceString(EmplrAcct, " ", "")
'  EmplrAcct = RightJustify(CDbl(EmplrAcct), 7)
'  RemitNumb = ReplaceString(RemitNumb, "-", "")
'  RemitNumb = ReplaceString(RemitNumb, " ", "")
'  RemitNumb = RightJustify(CDbl(RemitNumb), 6)
'
'  Header = "N" + QPTrim$(EmplrAcct$) + ReportingQtr + Year + Month1Employment
'  Header = Header + Month2Employment + Month3Employment + QtrTotalWages
'  Header = Header + QtrExcessWages + QtrTaxableWages + QPTrim$(RemitNumb$) + "E" + Blank12
'  Header = Header
'  Dim L As Integer
'  L = Len(Header)
'  OpenESCLineN ESCNHandle, NumOfESCRecs
'  L = Len(PrintLine(1))
'  ESC.Body = Header
'  ESC.CrLf = ESCRec(1).CrLf
'  Put ESCNHandle, 1, ESC
'  L = Len(ESC)
'  Close ESCNHandle
'  Call FixDuplicateEmployees
'  For x = 1 To PrintLineCnt
'    If x = 1 Then
'      Put #ESCFile, x, ESC
'    Else
'      Put #ESCFile, x, PrintLine(x - 1)
'    End If
'  Next x
  'Put #ESCFile, x, PrintLine(PrintLineCnt)
  'Close TRFile
  'Close
  
  Unload frmProcessing
  If chkSaveToFolder.Value = 1 Then
    MsgBox ("ESC file saved to " + ESCReport$ + ".")
  End If
  
  Exit Sub
  
SumESCData:
  For Cnt2 = 1 To 4  'put gross into correct quarter
    If (TransHRec(1).CheckDate >= QtrDates(Cnt2).LDate) And (TransHRec(1).CheckDate <= QtrDates(Cnt2).HDate) Then
      Select Case fpcomboPayType.Text
        Case "Social Security Wage"
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).SocGrossPay)
          If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
          End If
        Case "Gross Wage w/o SS Exempt Deductions"
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
          End If
          For x = 1 To NumOfDedRecs
            Get DedHandle, x, DedRec
              If QPTrim$(DedRec.DCSOC1) = "Y" Then
                EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) - TransHRec(1).DAmt(x))
              End If
          Next x
        Case "Gross Wage"
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
          End If
        Case "Retirement Wage"
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).RetGrossPay)
          If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
          End If
        Case Else
          MsgBox "Please make a selection in the Wage Preference list box."
          fpcomboPayType.SetFocus
          Exit Sub
        End Select
      Exit For
    End If
  Next
Return

SumAnnualESCData:
  WhichQtr = 0
  If (TransHRec(1).CheckDate >= YrStart) And (TransHRec(1).CheckDate <= YrEnd) Then
    If (TransHRec(1).CheckDate >= QtrDates(1).LDate) And (TransHRec(1).CheckDate <= QtrDates(1).HDate) Then
      WhichQtr = 1
      Select Case fpcomboPayType.Text
        Case "Social Security Wage"
          EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).SocGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).TaxFring)
          End If
        Case "Gross Wage w/o SS Exempt Deductions"
          EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).TaxFring)
          End If
          For x = 1 To NumOfDedRecs
            Get DedHandle, x, DedRec
              If QPTrim$(DedRec.DCSOC1) = "Y" Then
                EQtrWage2(1) = OldRound(EQtrWage2(1) - TransHRec(1).DAmt(x))
              End If
          Next x
        Case "Gross Wage"
          EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).TaxFring)
          End If
        Case "Retirement Wage"
          EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).RetGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(1) = OldRound(EQtrWage2(1) + TransHRec(1).TaxFring)
          End If
        Case Else
          MsgBox "Please make a selection in the Wage Preference list box."
          fpcomboPayType.SetFocus
          Exit Sub
        End Select
    ElseIf (TransHRec(1).CheckDate >= QtrDates(2).LDate) And (TransHRec(1).CheckDate <= QtrDates(2).HDate) Then
      WhichQtr = 2
      Select Case fpcomboPayType.Text
        Case "Social Security Wage"
          EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).SocGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).TaxFring)
          End If
        Case "Gross Wage w/o SS Exempt Deductions"
          EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).TaxFring)
          End If
          For x = 1 To NumOfDedRecs
            Get DedHandle, x, DedRec
              If QPTrim$(DedRec.DCSOC1) = "Y" Then
                EQtrWage2(2) = OldRound(EQtrWage2(2) - TransHRec(1).DAmt(x))
              End If
          Next x
        Case "Gross Wage"
          EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).TaxFring)
          End If
        Case "Retirement Wage"
          EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).RetGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(2) = OldRound(EQtrWage2(2) + TransHRec(1).TaxFring)
          End If
        Case Else
          MsgBox "Please make a selection in the Wage Preference list box."
          fpcomboPayType.SetFocus
          Exit Sub
        End Select
    ElseIf (TransHRec(1).CheckDate >= QtrDates(3).LDate) And (TransHRec(1).CheckDate <= QtrDates(3).HDate) Then
      WhichQtr = 3
      Select Case fpcomboPayType.Text
        Case "Social Security Wage"
          EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).SocGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).TaxFring)
          End If
        Case "Gross Wage w/o SS Exempt Deductions"
          EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).TaxFring)
          End If
          For x = 1 To NumOfDedRecs
            Get DedHandle, x, DedRec
              If QPTrim$(DedRec.DCSOC1) = "Y" Then
                EQtrWage2(3) = OldRound(EQtrWage2(3) - TransHRec(1).DAmt(x))
              End If
          Next x
        Case "Gross Wage"
          EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).TaxFring)
          End If
        Case "Retirement Wage"
          EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).RetGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(3) = OldRound(EQtrWage2(3) + TransHRec(1).TaxFring)
          End If
        Case Else
          MsgBox "Please make a selection in the Wage Preference list box."
          fpcomboPayType.SetFocus
          Exit Sub
        End Select
    ElseIf (TransHRec(1).CheckDate >= QtrDates(4).LDate) And (TransHRec(1).CheckDate <= QtrDates(4).HDate) Then
      WhichQtr = 4
      Select Case fpcomboPayType.Text
        Case "Social Security Wage"
          EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).SocGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).TaxFring)
          End If
        Case "Gross Wage w/o SS Exempt Deductions"
          EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).TaxFring)
          End If
          For x = 1 To NumOfDedRecs
            Get DedHandle, x, DedRec
              If QPTrim$(DedRec.DCSOC1) = "Y" Then
                EQtrWage2(4) = OldRound(EQtrWage2(4) - TransHRec(1).DAmt(x))
              End If
          Next x
        Case "Gross Wage"
          EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).GrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).TaxFring)
          End If
        Case "Retirement Wage"
          EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).RetGrossPay)
          If TransHRec(1).TaxFring > 0 Then
            EQtrWage2(4) = OldRound(EQtrWage2(4) + TransHRec(1).TaxFring)
          End If
        Case Else
          MsgBox "Please make a selection in the Wage Preference list box."
          fpcomboPayType.SetFocus
          Exit Sub
        End Select
    End If
  End If
Return


ErrorHandler:
  If NoADrive = True Then
    MsgBox "An error has occurred in accessing Drive A: and data has not been saved to disk."
  Else
    MsgBox "An error has occurred in the electronic file processing. Data has not been processed for electronic file."
  End If
  Close
  Unload frmProcessing
  Exit Sub
  
Return
 
PrintESCLine:
  If EQtrWage(CInt(fptxtQtr.Text)) = 0 Then
    GoTo SkipThisEmp
  End If
  
  EmpCnt = EmpCnt + 1
  ReDim ESCRec(1) As ESCMAG2DiskType
  
  CPos = InStr(Emp2Rec(1).EmpLName, ",")
  If CPos > 0 Then
    LName$ = Left$(Emp2Rec(1).EmpLName, CPos - 1)
  Else
    LName$ = Emp2Rec(1).EmpLName
  End If
  LName$ = QPTrim$(LName$)

  ESCRec(1).Blank1 = " "
  ESCRec(1).SSN = Emp2Rec(1).EmpSSN
  ESCRec(1).LastName = LName$
  ESCRec(1).EmpInitials = Left$(QPTrim$(Emp2Rec(1).EmpFName), 1) + Left$(QPTrim$(Emp2Rec(1).EmpLName), 1)
  ESCRec(1).EmpWages = RSet0(EQtrWage(Val(fptxtQtr.Text)), 9)

  If Left$(Emp2Rec(1).EMPSTATS, 1) = "S" Then
    ESCRec(1).SeasInd = "S"
  Else
    ESCRec(1).SeasInd = "N"
  End If
  ESCRec(1).RemitNumb = RemitNumb$
  ESCRec(1).EmplrAcct = EmplrAcct$

  ESCRec(1).BranchAcct = ""
  ESCRec(1).RQuarter = QPTrim(fptxtQtr.Text) 'Quarter$
  ESCRec(1).RYear = Year$
  ESCRec(1).EmplrName = UCase$(QPTrim$(Unit(1).UFEMPR))
  ESCRec(1).Blank2 = " "
  ESCRec(1).CrLf = Chr$(13) + Chr$(10)
  PrintLineCnt = PrintLineCnt + 1
  ReDim Preserve PrintLine(1 To PrintLineCnt) As ESCMAG2DiskType
  PrintLine(PrintLineCnt) = ESCRec(1)
  PrintLine(PrintLineCnt).LastName = PrintLine(PrintLineCnt).LastName
  PrintLine(PrintLineCnt).SSN = PrintLine(PrintLineCnt).SSN
  Put #ESCFile, , ESCRec(1)

SkipThisEmp:
Return

End Sub

Function RSet0$(Amt#, StrLen As Integer)
  Dim Temp$, NumStr$, Bit$
  Dim ChrPos As Integer, NewStr$
  Dim NumLen As Integer, StartPos As Integer
  Temp$ = String$(StrLen, "0")
  NumStr$ = QPTrim$(Str$(Amt#))

  Bit$ = Right$(NumStr$, 2)

  If InStr(Bit$, ".") Then
    NumStr$ = NumStr$ + "0"
  End If

  ChrPos = InStr(NumStr$, ".")
  If ChrPos Then
    NewStr$ = Left$(NumStr$, ChrPos - 1) + Mid$(NumStr$, ChrPos + 1)
  Else
    NewStr$ = NumStr$ + "00"
  End If

  NumStr$ = QPTrim$(NewStr$)

  NumLen = Len(NumStr$)
  StartPos = (StrLen - NumLen) + 1
  Mid$(Temp$, StartPos) = NumStr$
  RSet0$ = Temp$

End Function

Private Sub Form_Unload(Cancel As Integer)
  Call SaveESCMaxPay
End Sub

Private Sub fpcomboDiskFile_Change()
  If fpcomboDiskFile.Text = "y" Then
    fpcomboDiskFile.Text = "Y"
  End If
  If fpcomboDiskFile.Text = "n" Then
    fpcomboDiskFile.Text = "N"
  End If
  If fpcomboDiskFile.Text = "Y" Then
    chkSaveToFolder.Visible = True
  Else
    chkSaveToFolder.Visible = False
  End If
End Sub

Private Sub fpcomboDiskFile_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboDiskFile.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboDiskFile.ListIndex = -1
  End If
  If fpcomboDiskFile.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcomboPayType.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboPayType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPayType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPayType.ListIndex = -1
  End If
  If fpcomboPayType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtYear.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintText()
  Dim RptQuarter$, Qtr$, RptTitle$, RptName$
  Dim Year As String
  Dim CrLf2$, CrLf5$, CrLf6$, fmt$, Fmt2$, CrLf$, CrLf8$
  Dim MaxLines As Integer, IdxRecLen As Integer
  Dim UnitHandle As Integer, IdxFileSize&
  Dim EmpRecSize As Long, TRecSize As Long
  Dim NumOfRecs As Long, cnt As Integer
  Dim RHandle As Integer, THandle As Integer, DHandle As Integer
  Dim LineCnt As Integer
  Dim RecNo As Long, TransRecNum&, GrandTotal#
  Dim DoQtrLine As Boolean, FF$, PageNo As Integer
  Dim GrossOvr#, TotalGrossOvr#, Cnt2 As Integer
  Dim YTD2PrevQtr#, YTD2ThisQtr#, SubTotal#
  Dim EmpIdxLNameHandle As Integer, x As Integer
  Dim NameIdxLName As NameSortIdxType
  Dim Emp2Rec As EmpData2Type
  Dim DedRec As DedCodeRecType
  Dim NumOfDedRecs As Integer
  Dim DedHandle As Integer
  Dim ThisCnt As Integer
  
  If fptxtYear.Text = "" Then
     MsgBox "Please enter a Year"
     fptxtYear.SetFocus
     Exit Sub
  End If

  If Val(fptxtYear.Text) < 1920 Or Val(fptxtYear.Text) > 2099 Then
     MsgBox "Please enter a valid Year (####)"
     fptxtYear.SetFocus
     Exit Sub
  End If

  'the next if should never happen because the allowable
  'values in fptxtQtr are 1 2 3 & 4 only
  If Val(fptxtQtr.Text) < 1 Or Val(fptxtQtr.Text) > 4 Then
     MsgBox "Please enter a valid Quarter value"
     fptxtQtr.SetFocus
     Exit Sub
  End If
  
  RptQuarter$ = QPTrim$(fptxtQtr.Text)
  Year$ = QPTrim$(fptxtYear.Text)
  FF$ = Chr(12)
  CrLf2$ = CrLf$ + CrLf$
  CrLf5$ = CrLf2$ + CrLf2$ + CrLf$
  CrLf6$ = CrLf2$ + CrLf2$ + CrLf2$
  CrLf8$ = CrLf2$ + CrLf2$ + CrLf2$ + CrLf2$

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim GrsRpt(1 To 3) As ESCGrossWageRptType
  ReDim Line2(1) As String * 80

  ReDim EQtrWage(1 To 4) As Double
  ReDim QtrDates(1 To 4) As QtrDateType

  ReDim ToDisk1(1) As ESC2DiskRecType1
  ToDisk1(1).Seasonal = ""
  ToDisk1(1).Fill1 = ""
  ToDisk1(1).CrLf = CrLf$

  CrLf$ = Chr$(13) + Chr$(10)

  fmt$ = "#,###,##0.00"
  Fmt2$ = "###,###,##0.00"

  MaxLines = 25
  LineCnt = 0
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
    
  QtrDates(1).LDate = Date2Num("01/01/" + Year$)
  QtrDates(1).HDate = Date2Num("03/31/" + Year$)
  QtrDates(2).LDate = Date2Num("04/01/" + Year$)
  QtrDates(2).HDate = Date2Num("06/30/" + Year$)
  QtrDates(3).LDate = Date2Num("07/01/" + Year$)
  QtrDates(3).HDate = Date2Num("09/30/" + Year$)
  QtrDates(4).LDate = Date2Num("10/01/" + Year$)
  QtrDates(4).HDate = Date2Num("12/31/" + Year$)
  Qtr$ = QPTrim$(RptQuarter) + " " + Year$
  
  RptTitle$ = "ESC Quarterly Wage Report"
  
  If fpcomboDiskFile.Text = "Y" Then
    Call ESC2Disk
  End If
  
  RptName$ = "PRRPTS\ESCQTRT" + QPTrim$(RptQuarter) + ".RPT"
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) \ 2
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "ESC Quarterly Wage Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 7, RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  GoSub PrintESCHeader

  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If (Emp2Rec.LastTransRec <= 0) Or (Emp2Rec.ExcludeESC = "Y") Then
      GoTo SkipEm5
    End If
    TransRecNum& = Emp2Rec.LastTransRec
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      GoSub SumEmpESCData
      If TransHRec(1).PrevTransRec <= 0 Then
        GoSub PrintEmpESCLine
        If LineCnt >= MaxLines Then
          GoSub PrintSubTotals
          Print #RHandle, FF$
          GoSub PrintESCHeader
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm5:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
Next RecNo

  GoSub PrintESCGTotals
  Print #RHandle, FF$
  PageNo = PageNo + 1
  LSet Pg(1) = QPTrim$(Str$(PageNo))
  Print #RHandle, CrLf6$
  Print #RHandle, "Employer's Quarterly Tax and Wage Report Summary" + CrLf2$
  Print #RHandle, "  Total Wages:  "; Using$(Fmt2$, GrandTotal#) + CrLf$
  If Unit(1).ESCRTYPE = 2 Then
    Print #RHandle, " Excess Wages:  "; Using$(Fmt2$, TotalGrossOvr#) + CrLf$
    Print #RHandle, "Taxable Wages:  "; Using$(Fmt2$, OldRound(GrandTotal# - TotalGrossOvr#)) + CrLf$
  End If
  Print #RHandle, FF$
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True

  Close DHandle
  Close THandle
  RPTSetupPRN 123, RHandle '7/24
  Close RHandle
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  MainLog ("ESC Quarterly Wage Report processed.")
  
  Exit Sub

PrintESCHeader:
  LineCnt = 0
  PageNo = PageNo + 1
  LSet Pg(1) = QPTrim$(Str$(PageNo))
  '---
  Print #RHandle, CrLf6$
  '---
  Print #RHandle, "   " + QPTrim$(Unit(1).UFEMPR) + Space$(55) + Pg(1) + Space$(10) + QPTrim$(Unit(1).ESCEmplrNum) + CrLf$  'Unit(1).ESCRemitNum...Unit(1).ESCEmplrNum
  Print #RHandle, CrLf5$
  Print #RHandle, CrLf5$
  Print #RHandle, CrLf5$
  Print #RHandle, CrLf5$
  DoQtrLine = True
Return

SumEmpESCData:
  For Cnt2 = 1 To 4  'put gross into correct quarter
    If (TransHRec(1).CheckDate >= QtrDates(Cnt2).LDate) And (TransHRec(1).CheckDate <= QtrDates(Cnt2).HDate) Then
      Select Case fpcomboPayType.Text
      Case "Gross Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case "Gross Wage w/o SS Exempt Deductions"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).GrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        OpenDedCodeFile DedHandle
        NumOfDedRecs = LOF(DedHandle) / Len(DedRec)
        For x = 1 To NumOfDedRecs
          Get DedHandle, x, DedRec
          If QPTrim$(DedRec.DCSOC1) = "Y" Then
            EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) - TransHRec(1).DAmt(x))
          End If
        Next x
        Close DedHandle
        Exit For
      Case "Social Security Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).SocGrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case "Retirement Wage"
        EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).RetGrossPay) 'TransHRec(1).RetGrossPay
        If TransHRec(1).TaxFring > 0 Then 'added 7/22/03
          EQtrWage(Cnt2) = OldRound(EQtrWage(Cnt2) + TransHRec(1).TaxFring)
        End If
        Exit For
      Case Else
        MsgBox "Please make a selection in the Wage Preference list box."
        fpcomboPayType.SetFocus
        Exit Sub
      End Select
    End If
  Next
Return

PrintEmpESCLine:

  If Unit(1).ESCRTYPE = 2 Then
    If RptQuarter > 1 Then            'if Not 1st qtr, we will have to
      For Cnt2 = 1 To RptQuarter - 1  'examine gross thru prior qtr
        YTD2PrevQtr# = OldRound(YTD2PrevQtr# + EQtrWage(Cnt2))
      Next
      For Cnt2 = 1 To RptQuarter
        YTD2ThisQtr# = OldRound(YTD2ThisQtr# + EQtrWage(Cnt2))
      Next
      If YTD2PrevQtr# > Unit(1).TAXWBASE Then     'if the prior qtr was
        GrossOvr# = EQtrWage(RptQuarter)          'over also TaxBase also
      ElseIf YTD2ThisQtr# > Unit(1).TAXWBASE Then         'else if gross thru
        GrossOvr# = OldRound(YTD2ThisQtr# - Unit(1).TAXWBASE) 'this qtr is over
      Else   'still not over
        GrossOvr# = 0
      End If
'*** This deals with the 1ST quarter only!!
    ElseIf EQtrWage(1) > Unit(1).TAXWBASE Then  'else this is 1st qtr report
      GrossOvr# = OldRound(EQtrWage(1) - Unit(1).TAXWBASE)
    Else
      GrossOvr# = 0
    End If
'*** 1ST Quarter end
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter) 'OldRound(EQtrWage(RptQuarter) - GrossOvr#)
  Else       'not type 2 report
    GrsRpt(1).GrossPay = EQtrWage(RptQuarter)
  End If

  TotalGrossOvr# = OldRound(TotalGrossOvr# + GrossOvr#)

  If GrsRpt(1).GrossPay = 0 Then GoTo SkipEMPPrint

  RSet ToDisk1(1).GPay = Using$(fmt$, GrsRpt(1).GrossPay)
  LSet ToDisk1(1).ESSN = Left$(Emp2Rec.EmpSSN, 3) + "-" + Mid$(Emp2Rec.EmpSSN, 4, 2) + "-" + Mid$(Emp2Rec.EmpSSN, 6, 4)
  LSet ToDisk1(1).EName = Left$(Emp2Rec.EmpFName, 1) + " " + Left$(Emp2Rec.EmpLName, 18)

  If DoQtrLine Then
    RSet ToDisk1(1).Qtr = Qtr$
    DoQtrLine = False
  Else
    RSet ToDisk1(1).Qtr = " "
  End If
  ThisCnt = ThisCnt + 1
  Print #RHandle, ToDisk1(1).Qtr; ToDisk1(1).Fill1; ToDisk1(1).ESSN; ToDisk1(1).EName;
  Print #RHandle, ToDisk1(1).Seasonal; ToDisk1(1).GPay;
  Print #RHandle, CrLf$
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  SubTotal# = OldRound(SubTotal# + GrsRpt(1).GrossPay)
SkipEMPPrint:

  GrsRpt(1) = GrsRpt(2)
  YTD2PrevQtr# = 0
  YTD2ThisQtr# = 0
  For Cnt2 = 1 To 4
    EQtrWage(Cnt2) = 0
  Next

Return
PrintSubTotals:
  RSet Line2(1) = (Using$(fmt$, SubTotal#) + CrLf$)
  Print #RHandle, Line2(1)
  GrandTotal# = OldRound(GrandTotal# + SubTotal#)
  SubTotal# = 0
Return

PrintESCGTotals:
  If LineCnt < MaxLines Then
    For cnt = LineCnt To MaxLines - 1
      Print #RHandle, CrLf2$
    Next
  End If
  GoSub PrintSubTotals
  
  Return

DedExitRpt:

End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function GetMonth(ByVal TheDate As String) As Integer
  Dim x As Integer
  Dim Lgth As Integer
  Dim Mth As String
  Dim ch As String
  GetMonth = 0
  Lgth = Len(TheDate)
  For x = 1 To Lgth
    ch = Mid(TheDate, x, 1)
    If ch = "/" Then Exit For
    Mth = Mth + ch
  Next x
  
  Select Case Val(Mth)
    Case 1, 4, 7, 10
      GetMonth = 1
    Case 2, 5, 8, 11
      GetMonth = 2
    Case 3, 6, 9, 12
      GetMonth = 3
  End Select
  
End Function
Private Function RightJustify(ByVal Amt As Double, ByVal Lgth As Integer) As String
  RightJustify = ""
  Dim x As Integer
  Dim y As Integer
  Dim AmtStr As String
  Dim AmtLgth As Integer
  Dim ch As String
  Dim StopLgth As Integer
  Dim SendThis As String
  AmtStr = CStr(Amt)
  AmtStr = ReplaceString(AmtStr, ".", "")
  AmtLgth = Len(AmtStr)
  StopLgth = Lgth - AmtLgth + 1
  y = 1
  For x = 1 To Lgth
    If x >= StopLgth Then
      ch = Mid(AmtStr, y, 1)
      y = y + 1
      SendThis = SendThis + ch
    Else
      SendThis = SendThis + "0"
    End If
  Next x
  RightJustify = SendThis
  
End Function

Private Function LoadESCMaxPay()
  Dim AHandle As Integer
  Dim TextLine$
  LoadESCMaxPay = 0
  If Exist("PRData\ESCMaxSalary.dat") Then
    AHandle = FreeFile
    Open "PRData\ESCMaxSalary.dat" For Input As #AHandle
    Line Input #AHandle, TextLine
    LoadESCMaxPay = CDbl(TextLine)
'  Else
'    fptxtMaxSal.Text = "$0.00"
  End If
  Close AHandle

End Function

Private Sub SaveESCMaxPay()
  Dim AHandle As Integer
  Dim MaxSal As Double
  Dim MaxSalStr As String
  
  MaxSalStr = fptxtMaxSal.Text
  MaxSalStr = ReplaceString(MaxSalStr, "$", "")
  If IsNumeric(MaxSalStr) Then
    MaxSal = CDbl(MaxSalStr)
  End If
  
  If Exist("PRData\ESCMaxSalary.dat") Then
    Kill ("PRData\ESCMaxSalary.dat")
  End If
  
  AHandle = FreeFile
  Open "PRData\ESCMaxSalary.dat" For Output As AHandle
  Print #AHandle, MaxSal
  Close AHandle
End Sub

Private Sub fptxtMaxSal_LostFocus()
  Dim AmtVal As Double
  
  If QPTrim$(fptxtMaxSal.Text) = "" Then
    fptxtMaxSal.Text = "$0.00"
    Exit Sub
  End If
  AmtVal = CDbl(ReplaceString(fptxtMaxSal.Text, "$", ""))
  fptxtMaxSal.Text = FormatCurrency(AmtVal, 2, vbTrue)

End Sub

Private Sub FixDuplicateEmployees()
  Dim x As Integer
  Dim y As Integer
  Dim z As Integer
  Dim HoldAmt As Double
  Dim HoldStr As String
  Dim HoldESC As ESCMAG2DiskType
  Dim HoldIdx As Integer
  ReDim BadArr(1 To 1) As Integer
  Dim BadCnt As Integer
  Dim Testx As String
  Dim Testy As String
  
  For x = 1 To PrintLineCnt
    Testx = ReplaceString(PrintLine(x).SSN, "-", "")
    For y = 1 To PrintLineCnt
      Testy = ReplaceString(PrintLine(y).SSN, "-", "")
      If y = x Then GoTo SkipIt
      For z = 1 To BadCnt
        If BadArr(z) = y Or BadArr(z) = x Then
          GoTo SkipIt
        End If
      Next z
      'Debug.Print CStr(x) + "  " + CStr(y)
'      If x = 1 And y = 2 Then Stop
      If ReplaceString(PrintLine(x).SSN, "-", "") = ReplaceString(PrintLine(y).SSN, "-", "") And PrintLine(x).SeasInd = PrintLine(y).SeasInd Then
        HoldAmt = CDbl(PrintLine(x).EmpWages) + CDbl(PrintLine(y).EmpWages)
        PrintLine(x).EmpWages = RSet0(HoldAmt, 11)
        PrintLine(y).EmpWages = RSet0(0, 11)
        BadCnt = BadCnt + 1
        ReDim Preserve BadArr(1 To BadCnt) As Integer
        BadArr(BadCnt) = y
      End If
SkipIt:
    Next y
  Next x
  
  For x = 1 To BadCnt
    HoldIdx = BadArr(x)
    HoldESC = PrintLine(PrintLineCnt)
    PrintLine(PrintLineCnt).SSN = PrintLine(PrintLineCnt).SSN
    PrintLine(HoldIdx).SSN = PrintLine(HoldIdx).SSN
    PrintLine(BadArr(x)) = HoldESC
    PrintLineCnt = PrintLineCnt - 1
    ReDim Preserve PrintLine(1 To PrintLineCnt)
  Next x
'  For x = 1 To PrintLineCnt
'    Debug.Print PrintLine(x).LastName
'  Next x
End Sub
