VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPrintAdvanceLetter 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Advance Renewal Notice Letter"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11580
   Icon            =   "frmBLPrintAdvanceLetter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6324
      Left            =   1872
      TabIndex        =   3
      Top             =   1608
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11155
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
      Picture         =   "frmBLPrintAdvanceLetter.frx":08CA
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   375
         Left            =   3075
         TabIndex        =   2
         Tag             =   $"frmBLPrintAdvanceLetter.frx":08E6
         Top             =   3510
         Width           =   3705
         _Version        =   196608
         _ExtentX        =   6535
         _ExtentY        =   661
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         ColDesigner     =   "frmBLPrintAdvanceLetter.frx":0AB4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   3120
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Applications' menu."
         Top             =   4944
         Width           =   1896
         _Version        =   131072
         _ExtentX        =   3344
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
         ButtonDesigner  =   "frmBLPrintAdvanceLetter.frx":0DAB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5280
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAdvanceLetter.frx":0F89
         Top             =   4944
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
         ButtonDesigner  =   "frmBLPrintAdvanceLetter.frx":1064
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   360
         Left            =   4560
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAdvanceLetter.frx":1243
         Top             =   2160
         Width           =   1896
         _Version        =   131072
         _ExtentX        =   3344
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmBLPrintAdvanceLetter.frx":12FC
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   396
         Left            =   2640
         TabIndex        =   0
         Tag             =   $"frmBLPrintAdvanceLetter.frx":14E0
         Top             =   2160
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
      Begin EditLib.fpDateTime fptxtNewXDate 
         Height          =   348
         Left            =   2976
         TabIndex        =   1
         Tag             =   $"frmBLPrintAdvanceLetter.frx":15F6
         Top             =   2880
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
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   348
         Left            =   4752
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAdvanceLetter.frx":177A
         Top             =   2880
         Width           =   1932
         _Version        =   131072
         _ExtentX        =   3408
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
         ButtonDesigner  =   "frmBLPrintAdvanceLetter.frx":186A
      End
      Begin EditLib.fpText fptxtAdvLtrNum 
         Height          =   396
         Left            =   5136
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAdvanceLetter.frx":1A50
         Top             =   1248
         Width           =   492
         _Version        =   196608
         _ExtentX        =   868
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
         ControlType     =   1
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
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   636
         Left            =   672
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAdvanceLetter.frx":1B31
         Top             =   4944
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
         ButtonDesigner  =   "frmBLPrintAdvanceLetter.frx":1C01
      End
      Begin EditLib.fpDateTime fptxtPrintDate 
         Height          =   348
         Left            =   3744
         TabIndex        =   17
         Tag             =   "The date entered here will appear on the letters as the date you want as the date the letters were printed."
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Date:"
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
         Left            =   2208
         TabIndex        =   18
         Top             =   4128
         Width           =   1212
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation Range:"
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
         Left            =   912
         TabIndex        =   16
         Top             =   3552
         Width           =   2028
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
         Left            =   720
         TabIndex        =   14
         Top             =   5616
         Width           =   2100
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Using Advance Letter #:"
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
         Left            =   2304
         TabIndex        =   12
         Top             =   1344
         Width           =   2700
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Expiration Date:"
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
         Left            =   960
         TabIndex        =   9
         Top             =   2928
         Width           =   1884
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2844
         Left            =   480
         Top             =   1824
         Width           =   6876
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Advance Renewal Letter"
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
         Left            =   1776
         TabIndex        =   8
         Top             =   480
         Width           =   4332
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   336
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
         Left            =   1104
         TabIndex        =   7
         Top             =   2256
         Width           =   1356
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   528
      TabIndex        =   15
      Top             =   6762
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
      Height          =   6672
      Left            =   1728
      Top             =   1416
      Width           =   8100
   End
End
Attribute VB_Name = "frmBLPrintAdvanceLetter"
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

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdExit.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    fptxtNewXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    fptxtCatCode.ToolTipText = ""
    cmdCodeList.ToolTipText = ""
    fptxtAdvLtrNum.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdExit.ToolTipText = "Press to exit this screen."
'    cmdHelp.ToolTipText = "Press this button to bring up infoormational balloons for each field on this screen. Press this button again to deactivate the balloons."
'    cmdProcess.ToolTipText = "Press to begin the advance renewal letter printing process."
'    fptxtNewXDate.ToolTipText = "Advance renewal letters will be printed for all businesses whose business licenses will expire on the date you enter here."
'    fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'    fptxtCatCode.ToolTipText = "Enter the business license category ( or ALL) for which the advance renewal letter will be printed."
'    cmdCodeList.ToolTipText = "Press to bring up an interactive list of all available categories."
'    fptxtAdvLtrNum.ToolTipText = "This number refers to the advance letter style saved on the Town Setup screen. It is not editable."
    
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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
    Case vbKeyF4:
      SendKeys "%E"
      Call fpcmdXList_Click
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
      KillFile "advanceltrprint.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintAdvanceLetter.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim ThisZip$
  Dim NewYear$
  Dim One As Integer
  Dim DHandle As Integer
  
  lblBalloon.Visible = False
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdHelp.ToolTipText = "Press this button to bring up infoormational balloons for each field on this screen. Press this button again to deactivate the balloons."
'  cmdProcess.ToolTipText = "Press to begin the advance renewal letter printing process."
'  fptxtNewXDate.ToolTipText = "Advance renewal letters will be printed for all businesses whose business licenses will expire on the date you enter here."
'  fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'  fptxtCatCode.ToolTipText = "Enter the business license category ( or ALL) for which the advance renewal letter will be printed."
'  cmdCodeList.ToolTipText = "Press to bring up an interactive list of all available categories."
'  fptxtAdvLtrNum.ToolTipText = "This number refers to the advance letter style saved on the Town Setup screen. It is not editable."
  
  One = 1
  DHandle = FreeFile
  Open "advanceltrprint.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  fptxtAdvLtrNum.Text = QPTrim$(TownRec.LaserLtr)
  fptxtCatCode.Text = "ALL"
  fptxtNewXDate = Date
  NewYear = fptxtNewXDate.AdjustDate(fptxtNewXDate.DateValue, 1, 0, 0)
  fptxtNewXDate.DateValue = NewYear
  fpcmbRange.Text = "Up To And Include This Expiration"
  fpcmbRange.AddItem "Up To And Include This Expiration"
  fpcmbRange.AddItem "This Expiration Only"
  fptxtPrintDate = Date
End Sub
Private Sub cmdExit_Click()
  KillFile "advanceltrprint.dat"
  frmBLIssueAppsLics.Show
  DoEvents
  Unload frmBLPrintAdvanceLetter
End Sub

Private Sub cmdProcess_Click()
  If QPTrim$(fptxtNewXDate.Text) = "" Then
    fptxtNewXDate.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid expiration date."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
'    fptxtNewXDate.BackColor = &HFFFFFF
    fptxtNewXDate.SetFocus
    Exit Sub
  End If
  If CInt(Mid(fptxtNewXDate.Text, 7, 4)) > 2100 Then
    fptxtNewXDate.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid expiration date before 2100."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
'    fptxtNewXDate.BackColor = &HFFFFFF
    fptxtNewXDate.SetFocus
    Exit Sub
  ElseIf CInt(Mid(fptxtNewXDate.Text, 7, 4)) < 1980 Then
    fptxtNewXDate.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid expiration date after 1979."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
'    fptxtNewXDate.BackColor = &HFFFFFF
    fptxtNewXDate.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdvLtrNum.Text) = "1" Then
    Call PrintGraphics1
  ElseIf QPTrim$(fptxtAdvLtrNum.Text) = "2" Then
    Call PrintGraphics2
  Else
    Call PrintGraphics3
  End If
End Sub

Private Sub fpcmdXList_Click()
  frmBLXDateList.Show vbModal
End Sub

Private Sub fptxtCatCode_Change()
  If QPTrim$(fptxtCatCode.Text) = "" Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub

Private Sub PrintGraphics1()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType ' CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppCnt As Integer
  Dim dlm$, TotFee As Double
  Dim LaserRec1 As LaserLetterType1
  Dim LHandle As Integer
  Dim XDate As Integer
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim IssFee#
  Dim CatCode$, CustFee#
  Dim Snt&
  Dim RangeFlag As Integer
  Dim AddEmptyFields As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  AddEmptyFields = 0
  
  dlm = "~"
  
  RangeFlag = 2
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  Code$ = QPTrim$(fptxtCatCode.Text)
  Year$ = Mid(fptxtNewXDate.Text, 7, 4)

'  OpenSrchNameIdxFile IdxHandle
  OpenCustNameIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CustIdx)
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CustIdx
    IdxRecs(x) = CustIdx.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CHandle

  ReportFile$ = "BLRPTS\ARADVLTR.RPT"
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  If Exist("arlaser1.dat") Then
    OpenLaserFile1 LHandle
    Get LHandle, 1, LaserRec1
    Close LHandle
  Else
    frmBLMessageBoxJr.Label1.Caption = "Laser letter # 1 has been saved."
    frmBLMessageBoxJr.Label1.Height = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  XDate = Date2Num(fptxtNewXDate.Text)
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  
  frmBLShowPctComp.Label1 = "Loading Advance Letter #1"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For cnt = 1 To NumOfIdxRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo AllDoneHere
    Else
      If CustRec.VALID > XDate Then GoTo AllDoneHere
    End If
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo AllDoneHere
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT5)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm
      End If
    End If
    GoTo AllDoneHere
PrintForm:
      GoSub GetCustFee
      AddEmptyFields = 0
      '                              0
      Print #RptHandle, QPTrim$(LaserRec1.Header); dlm;
      '                             1                              2
      Print #RptHandle, QPTrim$(LaserRec1.TownOf); dlm; QPTrim$(LaserRec1.Address); dlm;
      '                                  3
      Print #RptHandle, QPTrim$(LaserRec1.CityStateZip); dlm;
      '                                4                         5                        6
      Print #RptHandle, QPTrim$(LaserRec1.Phone); dlm; QPTrim$(CustRec.ADDRESS2); dlm; QPTrim$(CustRec.BillName) + " #" + CStr(IdxRecs(cnt)); dlm;
      '                          7                      8
      Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm; RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
      
      For x = 0 To 11
        '9 - 20
        Print #RptHandle, QPTrim$(LaserRec1.Line1(x)); dlm;
      Next x
      TotFee = OldRound(FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + TownRec.IssFee)
      AppCnt = AppCnt + 1
      
      If QPTrim$(CustRec.BILLCAT1) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           21                              22                           23
        Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCodeDesc(CustRec.BILLCAT1); dlm; FeeAmt1#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT2) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           24                              25                          26
        Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCodeDesc(CustRec.BILLCAT2); dlm; FeeAmt2#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT3) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           27                              28                           29
        Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCodeDesc(CustRec.BILLCAT3); dlm; FeeAmt3#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT4) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           30                              31                           32
        Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCodeDesc(CustRec.BILLCAT4); dlm; FeeAmt4#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT5) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           33                              34                           35
        Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCodeDesc(CustRec.BILLCAT5); dlm; FeeAmt5#; dlm;
      End If
      
      For x = 1 To AddEmptyFields
        Print #RptHandle, ""; dlm;
      Next x
      '                    36             37
      Print #RptHandle, TotFee; dlm; TownRec.IssFee
      
AllDoneHere:
      
    frmBLShowPctComp.ShowPctComp cnt, NumOfIdxRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt

  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLAdvanceLtr.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Advance letter #1 processed.")
  
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
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C2:             'Catagory #2
  
  CustFee# = OldRound#(CustFee# + FeeAmt1#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt2#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt3#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt4#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
'  FeeAmt# = 0
  
Return


ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintAdvanceLetter", "PrintGraphics1", Erl)
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

Private Sub PrintGraphics2()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType ' CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppCnt As Integer
  Dim dlm$, TotFee As Double
  Dim LaserRec2 As LaserLetterType2
  Dim LHandle As Integer
  Dim XDate As Integer
  Dim Balance As Double
  Dim CustFee#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim IssFee#
  Dim CatCode$
  Dim Snt&
  Dim RangeFlag As Integer
  Dim AddEmptyFields As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  AddEmptyFields = 0
  
  dlm = "~"
  
  RangeFlag = 2
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  End If
  
  Balance = 0
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  Code$ = QPTrim$(fptxtCatCode.Text)
  Year$ = Mid(fptxtNewXDate.Text, 7, 4)

'  OpenSrchNameIdxFile IdxHandle
  OpenCustNameIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CustIdx)
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CustIdx
    IdxRecs(x) = CustIdx.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CHandle

  ReportFile$ = "BLRPTS\ARADVLT2.RPT"
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  If Exist("arlaser2.dat") Then
    OpenLaserFile2 LHandle
    Get LHandle, 1, LaserRec2
    Close LHandle
  Else
    frmBLMessageBoxJr.Label1.Caption = "Laser letter #2 has been saved."
    frmBLMessageBoxJr.Label1.Height = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  XDate = Date2Num(fptxtNewXDate.Text)
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  
  frmBLShowPctComp.Label1 = "Loading Advance Letter #2"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For cnt = 1 To NumOfIdxRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo AllDoneHere
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo AllDoneHere
    Else
      If CustRec.VALID > XDate Then GoTo AllDoneHere
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT5)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm
      End If
    End If
    GoTo AllDoneHere
PrintForm:
      GoSub GetCustFee
      
      '                             0                                   1
      Print #RptHandle, QPTrim$(LaserRec2.TownOf); dlm; QPTrim$(LaserRec2.Address); dlm;
      '                                  2                                 3
      Print #RptHandle, QPTrim$(LaserRec2.CityStateZip); dlm; QPTrim$(LaserRec2.Phone); dlm;
      '                                4                         5                       6                        7
      Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm; QPTrim$(CustRec.BillName); dlm; Date; dlm; QPTrim$(CustRec.CustNumb); dlm;
      '                          8                                              9
      Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm; RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
      
      For x = 0 To 7
        '10 - 17
        Print #RptHandle, QPTrim$(LaserRec2.Line1(x)); dlm;
      Next x
      
      TotFee = FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + TownRec.IssFee
      
      Balance = 0
      Balance = OldRound(CustRec.AcctBal + TotFee)
      
      AppCnt = AppCnt + 1
      AddEmptyFields = 0
      
      If QPTrim$(CustRec.BILLCAT1) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           18                              19                           20
        Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCodeDesc(CustRec.BILLCAT1); dlm; FeeAmt1#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT2) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           21                              22                           23
        Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCodeDesc(CustRec.BILLCAT2); dlm; FeeAmt2#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT3) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           24                              25                           26
        Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCodeDesc(CustRec.BILLCAT3); dlm; FeeAmt3#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT4) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           27                              28                           29
        Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCodeDesc(CustRec.BILLCAT4); dlm; FeeAmt4#; dlm;
      End If

      If QPTrim$(CustRec.BILLCAT5) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           30                              31                           32
        Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCodeDesc(CustRec.BILLCAT5); dlm; FeeAmt5#; dlm;
      End If
      
      For x = 1 To AddEmptyFields
        Print #RptHandle, ""; dlm;
      Next x
      '                   33
      Print #RptHandle, TotFee; dlm;
      If Balance = 0 Then
        '                 34       35
        Print #RptHandle, "0"; dlm; "0"; dlm;
      Else
      '                      34                     35
        Print #RptHandle, CStr(Balance); dlm; CustRec.AcctBal; dlm;
      End If
      '                       36
      Print #RptHandle, TownRec.IssFee

AllDoneHere:
      
    frmBLShowPctComp.ShowPctComp cnt, NumOfIdxRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt

  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  
  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLAdvLetter2.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Advance letter #2 processed.")
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
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C2:             'Catagory #2
  
  CustFee# = OldRound#(CustFee# + FeeAmt1#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt2#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt3#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt4#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
  
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintAdvanceLetter", "PrintGraphics2", Erl)
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
Private Sub fptxtNewXDate_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtNewXDate.BackColor = &HFFFFFF
End Sub
Private Sub fpcmbRange_Change()
  If QPTrim$(fpcmbRange.Text) = "" Then
    fpcmbRange.Text = "Up To And Include This Expiration"
  End If
End Sub

Private Sub fpcmbRange_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRange.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRange.ListIndex = -1
  End If
  If fpcmbRange.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPrintDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics3()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType ' CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppCnt As Integer
  Dim dlm$, TotFee As Double
  Dim LaserRec3 As LaserLetterType3
  Dim LHandle As Integer
  Dim XDate As Integer
  Dim RangeFlag As Integer
  Dim CustFee#, FeeAmt#
  Dim Prorate#
  Dim CatCode$, Snt&, Mult#
  Dim Revenue#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim XCnt As Integer
  Dim AddEmptyFields As Integer
  
  On Error GoTo ERRORSTUFF
  
  AddEmptyFields = 0
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  dlm = "~"
  
  RangeFlag = 2
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  Code$ = QPTrim$(fptxtCatCode.Text)
  Year$ = Mid(fptxtNewXDate.Text, 7, 4)

'  OpenSrchNameIdxFile IdxHandle
  OpenCustNameIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CustIdx)
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CustIdx
    IdxRecs(x) = CustIdx.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CHandle

  ReportFile$ = "BLRPTS\ARADVLT3.RPT"
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  If Exist("arlaser3.dat") Then
    OpenLaserFile3 LHandle
    Get LHandle, 1, LaserRec3
    Close LHandle
  Else
    frmBLMessageBoxJr.Label1.Caption = "Laser letter # 1 has been saved."
    frmBLMessageBoxJr.Label1.Height = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  XDate = Date2Num(fptxtNewXDate.Text)
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  
  frmBLShowPctComp.Label1 = "Loading Advance Letter #3"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For cnt = 1 To NumOfIdxRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo AllDoneHere
    Else
      If CustRec.VALID > XDate Then GoTo AllDoneHere
    End If
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo AllDoneHere
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      GoSub GetCustFee
      If Code$ = "ALL" Then
        GoSub PrintForm
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT5)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm
      End If
    End If
    GoTo AllDoneHere
PrintForm:
      AddEmptyFields = 0
      '                                0                                           1
      Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm; QPTrim$(CustRec.BillName) + " #" + CStr(IdxRecs(cnt)); dlm;
      '                             2                                 3
      Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm; RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
      
      For x = 0 To 5
        '4 - 9
        Print #RptHandle, QPTrim$(LaserRec3.Line1(x)); dlm;
      Next x
      
      For x = 0 To 3
        '10 - 13
        Print #RptHandle, QPTrim$(LaserRec3.Line2(x)); dlm;
      Next x
      
      AppCnt = AppCnt + 1
      TotFee = FeeAmt1 + FeeAmt2 + FeeAmt3 + FeeAmt4 + FeeAmt5 + TownRec.IssFee
      
      If QPTrim$(CustRec.BILLCAT1) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                           14                              15                           16
        Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCodeDesc(CustRec.BILLCAT1); dlm; FeeAmt1#; dlm;
      End If
      
      If QPTrim$(CustRec.BILLCAT2) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                            17                                 18                        19
        Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCodeDesc(CustRec.BILLCAT2); dlm; FeeAmt2#; dlm;
      End If
      
      If QPTrim$(CustRec.BILLCAT3) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                            20                                 21                        22
        Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCodeDesc(CustRec.BILLCAT3); dlm; FeeAmt3#; dlm;
      End If
      
      If QPTrim$(CustRec.BILLCAT4) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                            23                                 24                        25
        Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCodeDesc(CustRec.BILLCAT4); dlm; FeeAmt4#; dlm;
      End If
      
      If QPTrim$(CustRec.BILLCAT5) = "" Then
        AddEmptyFields = AddEmptyFields + 3
      Else
        '                            26                                 27                       28
        Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCodeDesc(CustRec.BILLCAT5); dlm; FeeAmt5#; dlm;
      End If
        
      For x = 1 To AddEmptyFields
        
        Print #RptHandle, ""; dlm;
      Next x
      '                   29                30
      Print #RptHandle, TotFee; dlm; TownRec.IssFee; dlm;
      
      '                      31                         32                                 33
      Print #RptHandle, fptxtPrintDate; dlm; QPTrim$(LaserRec3.Signer); dlm; QPTrim$(LaserRec3.Phone)
AllDoneHere:
    frmBLShowPctComp.ShowPctComp cnt, NumOfIdxRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt

  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLAdvanceLtr3.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Advance letter #3 processed.")
  
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
      End If    'End Test for Code
    Next Snt&
  End If        'End Test for Cat 1
  
  
C2:             'Category #2
  
  CustFee# = OldRound#(CustFee# + FeeAmt1#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt2#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt3#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt4#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
  
Return


ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintAdvanceLetter", "PrintGraphics3", Erl)
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

