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
      Height          =   6012
      Left            =   2112
      TabIndex        =   5
      Top             =   1398
      Width           =   7452
      _Version        =   196609
      _ExtentX        =   13144
      _ExtentY        =   10604
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
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmESCQrtRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboDiskFile 
         Height          =   405
         Left            =   5565
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   2925
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
         ColDesigner     =   "frmESCQrtRpt.frx":0C89
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3480
         TabIndex        =   4
         Top             =   4320
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
         ColDesigner     =   "frmESCQrtRpt.frx":102C
      End
      Begin EditLib.fpDateTime fptxtYear 
         Height          =   396
         Left            =   3792
         TabIndex        =   1
         Top             =   2208
         Width           =   1212
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
         Height          =   396
         Left            =   4752
         TabIndex        =   0
         Top             =   1536
         Width           =   636
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
         TabIndex        =   12
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
         ButtonDesigner  =   "frmESCQrtRpt.frx":13CF
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1248
         TabIndex        =   13
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
         ButtonDesigner  =   "frmESCQrtRpt.frx":15E6
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
         TabIndex        =   11
         Top             =   4410
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
         TabIndex        =   10
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
         Left            =   432
         TabIndex        =   9
         Top             =   3024
         Width           =   2028
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
         TabIndex        =   8
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
         Left            =   2016
         TabIndex        =   7
         Top             =   1680
         Width           =   2556
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
         Left            =   2592
         TabIndex        =   6
         Top             =   2304
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
  Private Temp_Class As Resize_Class

Private Sub cmdProcess_Click()
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
    Me.Visible = False
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
  
  Year$ = QPTrim$(fptxtYear.Text)
  
  frmProcessing.Label1.Caption = "Saving to Drive A:"
  DoEvents
  frmProcessing.Show , Me
  DoEvents

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim ESCRec(1) As ESCMAG2DiskType

  ReDim QtrDates(1 To 4) As QtrDateType

  QtrDates(1).LDate = Date2Num("01/01/" + Year$)
  QtrDates(1).HDate = Date2Num("03/31/" + Year$)
  QtrDates(2).LDate = Date2Num("04/01/" + Year$)
  QtrDates(2).HDate = Date2Num("06/30/" + Year$)
  QtrDates(3).LDate = Date2Num("07/01/" + Year$)
  QtrDates(3).HDate = Date2Num("09/30/" + Year$)
  QtrDates(4).LDate = Date2Num("10/01/" + Year$)
  QtrDates(4).HDate = Date2Num("12/31/" + Year$)

  
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

  RemitNumb$ = Unit(1).ESCRemitNum
  EmplrAcct$ = Unit(1).ESCEmplrNum
  
  If Len(RemitNumb$) = 0 Or QPTrim$(RemitNumb$) = "0" Then
    MsgBox "Please enter a ESC Remit Number in the Employer File located on the Control Maintenance Menu."
    Unload frmProcessing
    MsgBox "No data saved to drive A:"
    Close
    Exit Sub
  End If
  
  If Len(EmplrAcct$) = 0 Or QPTrim$(EmplrAcct$) = "0" Then
    MsgBox "Please enter a ESC Employer Number in the Employer File located on the Control Maintenance Menu."
    Unload frmProcessing
    MsgBox "No data saved to drive A:"
    Close
    Exit Sub
  End If
  
  On Local Error GoTo ErrorHandler
  ESCReport$ = "A:\ESCNC.WGS"
  RptFile = FreeFile
  Open ESCReport$ For Output As RptFile
  Close RptFile

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
   
      If TransHRec(1).PrevTransRec <= 0 Then
        GoSub PrintESCLine
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop

SkipEm:
  Next
  Close TRFile
  Close
  Unload frmProcessing
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

ErrorHandler:
  MsgBox "An error has occurred in accessing Drive A: and data has not been saved to disk."
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

Private Sub fpcomboDiskFile_Change()
  If fpcomboDiskFile.Text = "y" Then fpcomboDiskFile.Text = "Y"
  If fpcomboDiskFile.Text = "n" Then fpcomboDiskFile.Text = "N"
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

