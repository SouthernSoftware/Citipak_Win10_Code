VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPrintAppsRenwls 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Print Applications"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLPrintAppsRenwls.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6852
      Left            =   1920
      TabIndex        =   8
      Top             =   1002
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   12086
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLPrintAppsRenwls.frx":08CA
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   360
         Left            =   3120
         TabIndex        =   3
         Tag             =   $"frmBLPrintAppsRenwls.frx":08E6
         Top             =   4032
         Width           =   3708
         _Version        =   196608
         _ExtentX        =   6540
         _ExtentY        =   635
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
         ColDesigner     =   "frmBLPrintAppsRenwls.frx":0AA7
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   3120
         TabIndex        =   4
         Tag             =   $"frmBLPrintAppsRenwls.frx":0E0E
         Top             =   4656
         Width           =   3576
         _Version        =   196608
         _ExtentX        =   6308
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
         ColDesigner     =   "frmBLPrintAppsRenwls.frx":0F14
      End
      Begin LpLib.fpCombo fpcmbUseLogo 
         Height          =   348
         Left            =   5760
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAppsRenwls.frx":127B
         Top             =   2880
         Visible         =   0   'False
         Width           =   816
         _Version        =   196608
         _ExtentX        =   1439
         _ExtentY        =   614
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
         ColDesigner     =   "frmBLPrintAppsRenwls.frx":1381
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   624
         Left            =   3168
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Applications' menu."
         Top             =   5676
         Width           =   1908
         _Version        =   131072
         _ExtentX        =   3365
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
         ButtonDesigner  =   "frmBLPrintAppsRenwls.frx":16E8
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   630
         Left            =   5235
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAppsRenwls.frx":18C6
         Top             =   5670
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
         ButtonDesigner  =   "frmBLPrintAppsRenwls.frx":19D1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdCodeList 
         Height          =   408
         Left            =   4608
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAppsRenwls.frx":1BB0
         Top             =   2112
         Width           =   1848
         _Version        =   131072
         _ExtentX        =   3260
         _ExtentY        =   720
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
         ButtonDesigner  =   "frmBLPrintAppsRenwls.frx":1C69
      End
      Begin EditLib.fpText fptxtCatCode 
         Height          =   396
         Left            =   2688
         TabIndex        =   0
         Tag             =   $"frmBLPrintAppsRenwls.frx":1E4D
         Top             =   2112
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
      Begin EditLib.fpDateTime fptxtLicYear 
         Height          =   348
         Left            =   4224
         TabIndex        =   1
         Tag             =   $"frmBLPrintAppsRenwls.frx":1F63
         Top             =   2784
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
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
      Begin EditLib.fpText fptxtAppNum 
         Height          =   396
         Left            =   4464
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "This number refers to the application number saved from the Town Setup screen. It is not editable."
         Top             =   1200
         Width           =   732
         _Version        =   196608
         _ExtentX        =   1291
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
         Left            =   816
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAppsRenwls.frx":2186
         Top             =   5664
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
         ButtonDesigner  =   "frmBLPrintAppsRenwls.frx":2256
      End
      Begin EditLib.fpDateTime fptxtNewXDate 
         Height          =   348
         Left            =   3024
         TabIndex        =   2
         Tag             =   $"frmBLPrintAppsRenwls.frx":2439
         Top             =   3408
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
         Left            =   4800
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   $"frmBLPrintAppsRenwls.frx":25BD
         Top             =   3408
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
         ButtonDesigner  =   "frmBLPrintAppsRenwls.frx":26AD
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Town Logo?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5280
         TabIndex        =   22
         Top             =   2640
         Visible         =   0   'False
         Width           =   1788
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
         Left            =   960
         TabIndex        =   20
         Top             =   4080
         Width           =   2028
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date:"
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
         Left            =   1056
         TabIndex        =   19
         Top             =   3456
         Width           =   1836
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
         Left            =   864
         TabIndex        =   16
         Top             =   6336
         Width           =   2100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Application #:"
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
         Left            =   2688
         TabIndex        =   14
         Top             =   1296
         Width           =   1596
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "License Year:"
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
         Left            =   2448
         TabIndex        =   13
         Top             =   2832
         Width           =   1548
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ApplicationType:"
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
         TabIndex        =   12
         Top             =   4752
         Width           =   1788
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
         TabIndex        =   11
         Top             =   2208
         Width           =   1164
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Applications"
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
         TabIndex        =   10
         Top             =   480
         Width           =   4332
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3516
         Left            =   672
         Top             =   1824
         Width           =   6540
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   432
      TabIndex        =   17
      Top             =   6876
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
      MaxWidth        =   4000
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
      Height          =   7104
      Left            =   1800
      Top             =   882
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLPrintAppsRenwls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim PenAmt As Boolean

Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtAppNum.ToolTipText = ""
    fptxtCatCode.ToolTipText = ""
    cmdCodeList.ToolTipText = ""
    fptxtLicYear.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdHelp.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtAppNum.ToolTipText = "This number indicates the application form number selected from the Town Setup screen."
'    fptxtCatCode.ToolTipText = "Enter the desired category (or ALL) for which you wish to print business license applications."
'    cmdCodeList.ToolTipText = "Press to bring up  an interactive category listing."
'    fptxtLicYear.ToolTipText = "This date is the reference year from which the application will be printed. "
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print this report to a laser printer. Select text to print to a tractor fed printer."
'    cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
'    cmdHelp.ToolTipText = "Press "Turn Help On" to activate informational balloons that appear when you place the cursor over any field on this screen. Press 'Turn Help Off' to deactivate these informational balloons."
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintAppsRenwls.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim ThisZip$
  Dim NewYear$
  
  lblBalloon.Visible = False
'  fptxtAppNum.ToolTipText = "This number indicates the application form number selected from the Town Setup screen."
'  fptxtCatCode.ToolTipText = "Enter the desired category (or ALL) for which you wish to print business license applications."
'  cmdCodeList.ToolTipText = "Press to bring up  an interactive category listing."
'  fptxtLicYear.ToolTipText = "This date is the reference year from which the application will be printed. "
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print this report to a laser printer. Select text to print to a tractor fed printer."
'  cmdExit.ToolTipText = "Press 'Cancel' to exit this screen."
  PenAmt = False
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  If QPTrim$(TownRec.UseAmtPctYN) = "Amt" Then
    PenAmt = True
  End If
    
  One = 1
  DHandle = FreeFile
  Open "custappsRenews.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  fptxtNewXDate = Date
  NewYear = fptxtNewXDate.AdjustDate(fptxtNewXDate.DateValue, 1, 0, 0)
  fptxtNewXDate.DateValue = NewYear
  fptxtCatCode.Text = "ALL"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbRange.Text = "Up To And Include This Expiration"
  fpcmbRange.AddItem "Up To And Include This Expiration"
  fpcmbRange.AddItem "This Expiration Only"
  fptxtLicYear.Text = Mid(Date, 7, 4)
  fptxtAppNum.Text = TownRec.AppForm
  If InStr(TownRec.AppForm, "10") Then
    If Exist("townlogoadvltr3.bmp") Then
      fptxtLicYear.Enabled = False
      Label8.Visible = True
      fpcmbUseLogo.Visible = True
      fpcmbUseLogo.Text = "N"
      fpcmbUseLogo.AddItem "N"
      fpcmbUseLogo.AddItem "Y"
    End If
  End If
End Sub

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
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

Private Sub cmdExit_Click()
  Close
  frmBLIssueAppsLics.Show
  KillFile "custappsRenews.dat"
  DoEvents
  Unload frmBLPrintAppsRenwls
End Sub

Private Sub cmdProcess_Click()
  'get rid of old reprint file if any exists
  If Exist("artmpcus.dat") Then KillFile "artmpcus.dat"
  If fpcmbPrintOpt.Text = "Graphical" Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
    Call PrintGraphics
  ElseIf fpcmbPrintOpt.Text = "Text" Then
    frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
    Call PrintText
  Else
    Exit Sub
  End If
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

Private Sub fpcmdXList_Click()
  frmBLXDateList.Show vbModal
End Sub

Private Sub fptxtCatCode_Change()
  If QPTrim$(fptxtCatCode.Text) = "" Then
    fptxtCatCode.Text = "ALL"
  End If
End Sub

Private Sub PrintText()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$, FF$
  Dim AppFormat$
  Dim ReturnAdd$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType 'CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer, LCnt As Integer
  Dim cnt As Integer
  Dim ThisCode$, SCnt As Integer
  Dim TotalCust As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Lp As Integer
  Dim LicTotal#
  Dim CatCode$
  Dim ZZCnt As Integer
  Dim Snt&, Amt#
  Dim CODEDESC$
  Dim CodeType$
  Dim DESC1$
  Dim BaseAmt1#, BaseAmt2#, BaseAmt3#, BaseAmt4#, BaseAmt5#, BaseAmt6#
  Dim Revenue1#, Revenue2#, Revenue3#, Revenue4#, Revenue5#, Revenue6#
  Dim Percent1#, Percent2#, Percent3#, Percent4#, Percent5#, Percent6#
  Dim Maximum1#, Maximum2#, Maximum3#, Maximum4#, Maximum5#, Maximum6#
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim Nextcnt As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppType As Integer
  Dim TownLen As Integer
  Dim ThisTab As Integer
  Dim AddLen As Integer
  Dim CityLen As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim LessBase$
  Dim Dash$
  Dim TLen As Integer
  Dim TT$
  Dim BaseFee$
  Dim TotalFees As Double
  Dim AppCnt As Integer
  Dim MultiBY$
  Dim YrUpDown$(1 To 10)
  Dim CustFee#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim IssFee#
  Dim XDate As Integer
  Dim RangeFlag As Integer
  
  On Error GoTo ERRORSTUFF
  
  RangeFlag = 2
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  End If
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  IssFee# = TownRec.IssFee
  AppType = TownRec.AppForm
  XDate = Date2Num(fptxtNewXDate)
  Nextcnt = 1
  KillFile BLTempCustRecName 'kill temporary file used for reprints

  Code$ = QPTrim$(fptxtCatCode.Text)
  Year$ = fptxtLicYear.Text

'  OpenSrchNameIdxFile IdxHandle
  OpenCustNameIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(CustIdx)
  If NumOfIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CustIdx
    IdxRecs(x) = CustIdx.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CHandle

  OpenTempCustRec TempHandle
  
  ReportFile$ = "CUSTAPPS.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  If AppType > 1 Then
    OpenCatCodeFile CodeHandle
    NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
    If AppType = 2 Then
      GoSub PrintCustom2
    ElseIf AppType = 3 Then
      GoSub PrintCustom3
    ElseIf AppType = 4 Then
      GoSub PrintCustom4
    ElseIf AppType = 5 Then
      GoSub PrintCustom5
    ElseIf AppType = 6 Then
      GoSub PrintCustom6
    ElseIf AppType = 7 Then
      GoSub PrintCustom7
    ElseIf AppType = 8 Then
      GoSub PrintCustom8
    ElseIf AppType = 9 Then
      GoSub PrintCustom9
    End If
  Else
    GoSub PrintStandard
  End If

PrintCustom2: 'PrintText exits in this GoTo
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  Nextcnt = 1
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow2
    Else
      If CustRec.VALID > XDate Then GoTo NotNow2
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm2
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm2
      End If
    End If
  GoTo NotNow2
PrintForm2:
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); CStr(Nextcnt); Tab(30); "LICENSE APPLICATION"
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf); Tab(58); "ACCOUNT NO.    " + Using("####0", IdxRecs(cnt));
  Print #RptHandle, Tab(55); "START DATE: "; Tab(67); UCase(QPTrim$(TownRec.AppStartMonth)) + " " + CStr(TownRec.AppStartDay) + ", " + YrUpDown(1)
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); "APPLICANT'S NAME:    "; QPTrim$(CustRec.BillName)
  Print #RptHandle, Tab(5); "APPLICANT'S ADDRESS: "; QPTrim$(CustRec.ADDRESS1)
  Print #RptHandle, Tab(5); "                     "; QPTrim$(CustRec.ADDRESS2)
  Print #RptHandle, Tab(5); "                     "; QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + "  " + QPTrim$(CustRec.ZipCode)
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); "TAX MAP______  BLOCK______ LOT______     ZONING DISTRICT_________________"
  Print #RptHandle, Tab(5); "FEDERAL ID/SS NUMBER__________________ " + QPTrim$(TownRec.AppState) + " TAX ID NUMBER__________________"
  Print #RptHandle, Tab(5); "TYPE OF BUSINESS:________________________________________________________"
  Print #RptHandle, Tab(5); "APPLICATION FOR:___ NEW___ RENEWAL___ GOING OUT OF BUSINESS(DATE)________"
  Print #RptHandle, Tab(5); "OWNERSHIP:___ CORPORATION___ PARTNERSHIP____ INDIVIDUAL-NO EMPLOYEES_____"
  Print #RptHandle, Tab(5); "NAME OF OWNER, PARTNER OR PRINCIPAL______________________________________"
  Print #RptHandle, Tab(5); "TELEPHONE NO. LOCAL:_____________ HOME:____________ EMERGENCY:___________"
  Print #RptHandle, Tab(5); "FAX NO._____________  E-MAIL:____________________________________________"
  Print #RptHandle,
  Print #RptHandle, Tab(5); "IS HAZARDOUS WASTE INVOLVED IN OPERATION? ____NO ____YES (ATTACH DETAILS)"
  Print #RptHandle, Tab(5); "CODE CLEARANCE: __ZONING ___INSPECTION __FIRE __HEALTH ___LAW ENFORCEMENT"
  Print #RptHandle,
  Print #RptHandle, Tab(28); "COMPUTATION OF LICENSE TAX"
  Print #RptHandle, Tab(5); "COMPUTE LICENSE TAX ACCORDING TO THE FOLLOWING SCHEDULE AND MAKE CHECKS"
  Print #RptHandle, Tab(5); "PAYABLE TO: "; QPTrim$(TownRec.AppTownOf) + ". DELIVER BY DUE DATE: "; Tab(61); UCase(QPTrim$(TownRec.AppLicRetMonth)) + " " + CStr(TownRec.AppLicRetDay) + ", " + YrUpDown(2)
  Print #RptHandle,
  Print #RptHandle, Tab(5); "GROSS INCOME FOR PRECEDING CALENDAR OR FISCAL YEAR....$_________________"
  Print #RptHandle, Tab(5); "LESS INCOME ON WHICH A LICENSE TAX WAS PAID TO ANOTHER"
  Print #RptHandle, Tab(5); "CITY OR COUNTY FOR OPERATIONS OUTSIDE CITY/COUNTY.....$_________________"
  Print #RptHandle, Tab(5); "BALANCE OF GROSS INCOME SUBJECT TO LICENSE TAX........$_________________"
  Print #RptHandle, Tab(5); "TAX:   RATE CLASS MINIMUM ON FIRST " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(1))) + ": " + QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(1))) + " PLUS"
  Print #RptHandle, Tab(5); QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(2))) + " PER " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(2))) + " FOR INCOME OVER " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(3)))
  Print #RptHandle, Tab(5); "[See declining rate schedule for over $1 million]       [OFFICE USE ONLY]"
  Print #RptHandle, Tab(5); "                           TOTAL LICENSE TAX $_________ [PAYMENT RECORD]"
  If PenAmt = False Then
    Print #RptHandle, Tab(5); "PENALTY AFTER DUE DATE IS " + CStr(TownRec.AppPct) + "% PER MONTH $_________ [CHECK NO. ____________]"
  Else
    Print #RptHandle, Tab(5); "PENALTY AFTER DUE DATE IS " + QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + " PER MONTH $_________ [CHECK NO. ____________]"
  End If
  Print #RptHandle, Tab(5); "TOTAL LICENSE TAX AND PENALTY $_________     [DATE RECEIVED____________]"
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(35); "CERTIFICATION"
  Print #RptHandle, Tab(5); "I (WE) DO CERTIFY THAT THE ABOVE INFORMATION AND AMOUNT RETURNED AS GROSS"
  Print #RptHandle, Tab(5); "INCOME FROM MY BUSINESS IS TRUE AND CORRECT. AND I HAVE MADE NO DEDUCTIONS"
  Print #RptHandle, Tab(5); "EXCEPT INCOME ON WHICH I HAVE PAID BUSINESS LICENSE TAX TO ANOTHER CITY OR"
  Print #RptHandle, Tab(5); "COUNTY, FOR WHICH I HAVE PROOF OF PAYMENT. I AM FAMILIAR WITH THE PENALTY"
  Print #RptHandle, Tab(5); "PROVISIONS OF THE ORDINANCE AND GROUNDS FOR LICENSE REVOCATION, INCLUDING"
  Print #RptHandle, Tab(5); "MAKING FALSE OR FRAUDULENT STATEMENTS IN THIS APPLICATION. I CERTIFY THAT"
  Print #RptHandle, Tab(5); "ALL BUSINESS PERSONAL PROPERTY TAXES AND PAYABLES DUE TO THE CITY/COUNTY"
  Print #RptHandle, Tab(5); "HAVE BEEN PAID, AND THAT THE ABOVE BUSINESS NAME IS THE SAME AS REPORTED"
  Print #RptHandle, Tab(5); "ON DOCUMENTS FILED WITH THE STATE AND FEDERAL GOVERNMENTS. I UNDERSTAND MY"
  Print #RptHandle, Tab(5); "BUSINESS INCOME TAX RETURNS AND OTHER DOCUMENTS MAY BE INSPECTED TO VERIFY"
  Print #RptHandle, Tab(5); "GROSS INCOME OR OTHER BUSINESS DATA."
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); "___________________________________________________________________________"
  Print #RptHandle, Tab(5); "SIGNATURE                          TITLE                           DATE"
  Print #RptHandle, Chr$(12);
  AppCnt = AppCnt + 1
  TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
  TempCustRec.AppType = AppType
  TempCustRec.ThisYear = Year$
  TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
  For x = 1 To 5
    TempCustRec.Fee(x) = 0
    TempCustRec.CatCode(x) = ""
    TempCustRec.CatDesc(x) = ""
  Next x
  TempCustRec.MiscNum = 0
  TempCustRec.IssFee = IssFee#
  Put TempHandle, Nextcnt, TempCustRec
  Nextcnt = Nextcnt + 1
NotNow2:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #2 processed in text format.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom3:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  Nextcnt = 1
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow3
    Else
      If CustRec.VALID > XDate Then GoTo NotNow3
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm3
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm3
      End If
    End If
    GoTo NotNow3
PrintForm3:
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); CStr(Nextcnt); Tab(ThisTab); QPTrim$(TownRec.AppTownOf) '"TOWN OF RIVERSIDE"
    Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Business Name: "; QPTrim$(CustRec.CustName); Tab(50); QPTrim$(CustRec.CustNumb)
    Print #RptHandle, Tab(5); "              -----------------------------------------------------------"
    Print #RptHandle, Tab(5); "Street Address of Business: "
    Print #RptHandle, Tab(5); "                           ----------------------------------------------"
    Print #RptHandle, Tab(5); "Zoning of Business Location: "
    Print #RptHandle, Tab(5); "                            ---------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, Tab(5); "Applicant's Name: "; QPTrim$(CustRec.BillName)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "Applicant's Address: "; QPTrim$(CustRec.ADDRESS1)
    Print #RptHandle, Tab(5); "                    -----------------------------------------------------"
    If QPTrim$(CustRec.WPHONE) = "(" Then CustRec.WPHONE = ""
    Print #RptHandle, Tab(23); QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); Tab(57); "Phone: "; QPTrim$(CustRec.WPHONE)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Rem 22 lines printed here
    Print #RptHandle, Tab(5); "TYPE OF BUSINESS LICENSE APPLYING FOR:"
    Print #RptHandle, Tab(5); ""
    If TownRec.IssFee > 0 Then
      Print #RptHandle, Tab(5); "_______ Contracting or Construction " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + " plus " + QPTrim(Using("$#,##0.00", TownRec.IssFee)) + " Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " plus " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + " of " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + " of gross receipts"
      Print #RptHandle, Tab(5); "        over " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + " plus " + QPTrim$(Using("$#, ##0.00", TownRec.IssFee)) + " Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + " plus " + QPTrim(Using("$#,##0.00", TownRec.IssFee))
      Print #RptHandle, Tab(5); "        Issuance Fee."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Repair, Personal, Business or Delivery Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + " plus " + QPTrim(Using("$#,##0.00", TownRec.IssFee))
      Print #RptHandle, Tab(5);
    Else
      Print #RptHandle, Tab(5); "_______ Contracting or Construction: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + "."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " plus " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + " of " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + " of gross receipts"
      Print #RptHandle, Tab(5); "        over " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + "."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + "."
      Print #RptHandle, Tab(5);
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "_______ Repair, Personal, Business or Delivery Service: " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + "."
      Print #RptHandle, Tab(5);
    End If
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Other (Specify) ______________________________________________"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "Estimate of ______________ gross receipts or preceding year's gross "
    Print #RptHandle, Tab(5); "receipts ______________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "AMOUNT OF LICENSE TAX FOR " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", THROUGH " + QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim$(YrUpDown(2)) + " IS:$_______"
    Print #RptHandle, Tab(5); "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "ACTIVITY SHALL BE CONDUCTED: ____________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "I certify that the statements and figures set forth on this application"
    Print #RptHandle, Tab(5); "are true to the best of my knowledge."
    Print #RptHandle, Tab(5); "                                      ___________________________________"
    Print #RptHandle, Tab(5); "                                            Signature of Applicant"
    Print #RptHandle, Tab(5); ""
    If PenAmt = False Then
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    Else
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    End If
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "Return Application and Fee to:"
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppAdd1)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    Print #RptHandle, Chr$(12);
    AppCnt = AppCnt + 1
    TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1
NotNow3:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  MainLog ("Application #3 processed in text format.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom4:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(4)) = "Curr" Then
    YrUpDown(4) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "+1" Then
    YrUpDown(4) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "-1" Then
    YrUpDown(4) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(5)) = "Curr" Then
    YrUpDown(5) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "+1" Then
    YrUpDown(5) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "-1" Then
    YrUpDown(5) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(6)) = "Curr" Then
    YrUpDown(6) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "+1" Then
    YrUpDown(6) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "-1" Then
    YrUpDown(6) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(7)) = "Curr" Then
    YrUpDown(7) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "+1" Then
    YrUpDown(7) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "-1" Then
    YrUpDown(7) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(8)) = "Curr" Then
    YrUpDown(8) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "+1" Then
    YrUpDown(8) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "-1" Then
    YrUpDown(8) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(9)) = "Curr" Then
    YrUpDown(9) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "+1" Then
    YrUpDown(9) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "-1" Then
    YrUpDown(9) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(10)) = "Curr" Then
    YrUpDown(10) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "+1" Then
    YrUpDown(10) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "-1" Then
    YrUpDown(10) = CStr(CInt(Year$) - 1)
  End If
  TownLen = Len(QPTrim$(TownRec.AppTownOf))

  ThisTab = TownLen / 2 'start centering process
  ThisTab = Abs(39 - ThisTab) '63 = end of line 2...63 - 16 = 47 length of line from beginning tab
  '...47/2 = 23.5 ...+ 16 = middle point of line...round down to 39

  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))

  tab2 = TownLen / 2
  tab2 = Abs(38 - tab2) '38 = mid point of line 3
  Tab3 = AddLen / 2
  Tab3 = Abs(38 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(38 - Tab4)
  Nextcnt = 1
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow4
    Else
      If CustRec.VALID > XDate Then GoTo NotNow4
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm4
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm4
      End If
    End If
    GoTo NotNow4
PrintForm4:
    Print #RptHandle, "" '33
    Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(tab2); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE" 'line 2
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1)); Tab(70); "PAGE 1"
    Print #RptHandle, ""
    Print #RptHandle, Tab(2); "Dear Business Owner:"
    Print #RptHandle,
    Print #RptHandle, Tab(2); "     For the purpose of computing Business, Professional and Occupational"
    Print #RptHandle, Tab(2); "License (BPOL) Tax promulgated by Virginia Code Section 58.1-3700 et seq."
    Print #RptHandle, Tab(2); "and " + QPTrim$(TownRec.AppCity) + " Town Ordinance #" + QPTrim$(TownRec.AppCityOrd) + " adopted "
    Print #RptHandle, Tab(2); MakeRegDate(TownRec.AppAdoptDate) + " please complete and return this form with the required"
    Print #RptHandle, Tab(2); "information no later than " + QPTrim$(TownRec.AppDiscMonth) + " " + CStr(TownRec.AppDiscDay) + ", " + QPTrim$(YrUpDown(2)) + "."
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(2); "Respectfully,"
    Print #RptHandle, Tab(2); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(2); QPTrim$(TownRec.AppMayorCouncil)
    Print #RptHandle, Tab(2); String$(76, "-")
    Print #RptHandle, Tab(2);
    Print #RptHandle, Tab(2);
    Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(Tab3); QPTrim$(TownRec.TownAdd1)
    Print #RptHandle, Tab(Tab4); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    Print #RptHandle, Tab(24); "Application for Town Licenses" 'line 3
    Print #RptHandle,
    Print #RptHandle, Tab(2); "For period beginning " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", " + QPTrim$(YrUpDown(3)) + " (or start of business in " + QPTrim$(YrUpDown(4)) + ")"
    Print #RptHandle, Tab(2); "and ending " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(5))
    Print #RptHandle, Tab(2);
    Print #RptHandle, Tab(2); "NAME OF APPLICANT: "; QPTrim$(CustRec.BillName)
    Print #RptHandle, Tab(2); "       TRADING AS: "; QPTrim$(CustRec.CustName)
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "BUSINESS ADDRESS:"; Tab(40); "HOME ADDRESS"
    Print #RptHandle, Tab(2); "MAIL: "; QPTrim$(CustRec.ADDRESS1); Tab(40); "MAIL: ________________________________"
    Print #RptHandle, Tab(8); QPTrim$(CustRec.ADDRESS2)
    Print #RptHandle, Tab(8); RTrim$(CustRec.City); " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); Tab(40); "      ________________________________"
    Print #RptHandle, Tab(2); "911:  ______________________________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(8); "______________________________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "PHONE: _________________________"; Tab(40); "PHONE: ______________________________"
    Print #RptHandle,
    Print #RptHandle, Tab(2); "A SEPARATE LICENSE WILL BE ISSUED FOR EACH TYPE OF BUSINESS"
    Print #RptHandle, Tab(2); "PERFORMED, AS REQUIRED PER THE " + UCase(QPTrim$(TownRec.AppCityOrd)) + ".  THIS WILL NOT"
    Print #RptHandle, Tab(2); "RESULT IN ANY ADDITONAL COST TO BUSINESSES.  PLEASE REPORT GROSS"
    Print #RptHandle, Tab(2); "RECEIPTS FOR EACH CLASSIFICATION THAT APPLIES TO YOUR BUSINESS."
    Print #RptHandle, Chr$(12);
    Print #RptHandle, Tab(ThisTab - 1); QPTrim$(TownRec.TownName)
    Print #RptHandle, Tab(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
    Print #RptHandle, Tab(31); "For Year: " + QPTrim$(YrUpDown(1)); Tab(70); "PAGE 2"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "WHOLESALE MERCHANT:"
    Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppWholeMonth) + "-" + CStr(TownRec.AppWholeDay) + "-" + QPTrim(YrUpDown(6)) + " as shown by applicants records"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); Tab(60); "$_______________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "RETAIL MERCHANT:"
    Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppRetailMonth) + "-" + CStr(TownRec.AppRetailDay) + "-" + QPTrim(YrUpDown(7)) + " as shown by applicants records"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); Tab(60); "$_______________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "FINANCIAL, REAL ESTATE AND PROFESSIONAL:"
    Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppFinMonth) + "-" + CStr(TownRec.AppFinDay) + "-" + QPTrim(YrUpDown(8)) + " as shown by applicants records"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); Tab(60); "$_______________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "CONTRACTING:"
    Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppContMonth) + "-" + CStr(TownRec.AppContDay) + "-" + QPTrim(YrUpDown(9)) + " as shown by applicants records"
    Print #RptHandle, Tab(2); "(Subject to Virginia Code Sec 58.1-3715)"
    Print #RptHandle, Tab(2); Tab(60); "$_______________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "REPAIR, PERSONAL or BUSINESS SERVICES:"
    Print #RptHandle, Tab(2); "Gross Receipts through " + CStr(TownRec.AppRepairMonth) + "-" + CStr(TownRec.AppRepairDay) + "-" + QPTrim(YrUpDown(10)) + " as shown by applicants records"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); Tab(60); "$_______________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "If uncertain of your business classification(s), please call the Town Office at"
    Print #RptHandle, Tab(2); QPTrim$(TownRec.AppPhone) + " for assistance."
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "I do affirm that the foregoing figures are true, complete and accurate to the"
    Print #RptHandle, Tab(2); "best of my knowledge."
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(35); "Signature ___________________________________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(34); "Print Name ___________________________________"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(24); "*** IMPORTANT ***"
    Print #RptHandle, Tab(2); ""
    Print #RptHandle, Tab(2); "APPLICATION MUST BE RETURNED PRIOR TO " + UCase(QPTrim$(TownRec.AppFiscMonth)) + " " + CStr(TownRec.AppFiscDay) + " OF EACH YEAR"
    Print #RptHandle, Tab(2); "TO AVOID PENALTY. LICENSE FEES ARE DUE PRIOR TO " + UCase(QPTrim$(TownRec.AppLicRetMonth)) + " " + CStr(TownRec.AppLicRetDay)
    Print #RptHandle, Tab(2); "OF EACH YEAR TO AVOID PENALTY AND INTEREST. INTENTIONALLY PROVIDING"
    Print #RptHandle, Tab(2); "INSUFFICIENT OR INACCURATE INFORMATION MAY RESULT IN LEGAL RECOURSE"
    Print #RptHandle, Tab(2); "BY THE TOWN OF " + UCase(QPTrim$(TownRec.AppCity)) + " AS SET FORTH BY VIRGINIA CODE."
    Print #RptHandle, Tab(2); Chr$(12);
    AppCnt = AppCnt + 1
    TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1
NotNow4:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #4 processed in text format.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom5:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  AppCnt = 0
  Nextcnt = 1
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(38 - ThisTab)
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow5
    Else
      If CustRec.VALID > XDate Then GoTo NotNow5
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm5
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm5
      End If
    End If
    GoTo NotNow5
PrintForm5:
    AppCnt = AppCnt + 1
    Print #RptHandle, ""
    Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(ThisTab); QPTrim$(TownRec.TownName)
    Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Business Name: "; QPTrim$(CustRec.CustName)
    Print #RptHandle, Tab(5); "              -----------------------------------------------------------"
    Print #RptHandle, Tab(5); "Street Address of Business: "
    Print #RptHandle, Tab(5); "                           ----------------------------------------------"
    Print #RptHandle, Tab(5); "Zoning of Business Location: "
    Print #RptHandle, Tab(5); "                            ---------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, Tab(5); "Applicant's Name: "; QPTrim$(CustRec.BillName)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Print #RptHandle, Tab(5); "Applicant's Address: "; QPTrim$(CustRec.ADDRESS1)
    Print #RptHandle, Tab(5); "                    -----------------------------------------------------"
    Print #RptHandle, Tab(5); "Telephone Number: "; QPTrim$(CustRec.WPHONE)
    Print #RptHandle, Tab(5); "                 --------------------------------------------------------"
    Rem 22 lines printed here
    Print #RptHandle, Tab(5); "TYPE OF BUSINESS LICENSE APPLYING FOR:"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Contracting or Construction " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + " or " + Using("#.###", TownRec.AppCentsPer(1)) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1)))
    Print #RptHandle, Tab(5); "           gross receipts whichever is greater."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_______ Retail Sales " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + " or " + Using("#.###", TownRec.AppCentsPer(2)) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(2))) + " whichever"
    Print #RptHandle, Tab(5); "           is greater."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Financial, Real Estate or Professional Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + " or "
    Print #RptHandle, Tab(5); "           " + Using("#.###", TownRec.AppCentsPer(3)) + " cents per " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(3))) + " whichever is greater."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_______ Repair, Personal or Business Service " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + " or " + Using("#.###", TownRec.AppCentsPer(4)) + " cents per "
    Print #RptHandle, Tab(5); "           " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(4))) + " whichever is greater."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "_______ Other (Specify) ______________________________________________"
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); "Estimate of ______________ gross receipts or preceding year's gross "
    Print #RptHandle, Tab(5); "receipts ______________________. Enclose copy of most recent schedule C"
    Print #RptHandle, Tab(5); "or other comparable federal document."
    Print #RptHandle, Tab(5); "AMOUNT OF LICENSE TAX FOR " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", THROUGH " + QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim(YrUpDown(2)) + " IS:$_______"
    Print #RptHandle, Tab(5); "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "ACTIVITY SHALL BE CONDUCTED: ____________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "_________________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "I certify that the statements and figures set forth on this application"
    Print #RptHandle, Tab(5); "are true to the best of my knowledge."
    Print #RptHandle, Tab(5); "                                      ___________________________________"
    Print #RptHandle, Tab(5); "                                            Signature of Applicant"
    Print #RptHandle, Tab(5); ""
    If PenAmt = False Then
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    Else
      Print #RptHandle, Tab(5); "To Avoid Late Penalty Charge of " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."
    End If
    Print #RptHandle, Tab(5);
    Print #RptHandle, Tab(5); "Return Application and Fee to:"
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppAdd1)
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    Print #RptHandle, Chr$(12);
    TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1
NotNow5:
  Next cnt
  Close         'Close all open files now
  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #5 processed in text format.")
  
  Exit Sub
'-----------------------------------------------------
PrintCustom6:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  Dash$ = String$(30, "_")
  MultiBY$ = CStr(TownRec.AppPct)
  Nextcnt = 1
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  ThisTab = TownLen / 2
  ThisTab = Abs(39 - ThisTab)
  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))
  tab2 = TownLen / 2
  tab2 = Abs(39 - tab2)
  Tab3 = AddLen / 2
  Tab3 = Abs(39 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(39 - Tab4)
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow6
    Else
      If CustRec.VALID > XDate Then GoTo NotNow6
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm6
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm6
      End If
    End If
    GoTo NotNow6
PrintForm6:
    Print #RptHandle, "" '
    Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(ThisTab); QPTrim$(TownRec.AppTownOf) '"TOWN OF ELLOREE"
    Print #RptHandle, Tab(Tab3); QPTrim$(TownRec.AppAdd1) '"P.O. BOX 28"
    Print #RptHandle, Tab(Tab4); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip) '"ELLOREE, S.C. 29047"
    Print #RptHandle, Tab(20); "APPLICATION FOR BUSINESS LICENSE FOR YEAR "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); CustRec.BillName
    Print #RptHandle, Tab(5); CustRec.ADDRESS1
    Print #RptHandle, Tab(5); CustRec.ADDRESS2
    Print #RptHandle, Tab(5); QPTrim$(CustRec.City); ", "; CustRec.State; " "; CustRec.ZipCode
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "To engage in business or profession, make a separate application"
    Print #RptHandle, Tab(5); "for each business and each location.  Send fee with application to"
    Print #RptHandle, Tab(5); "The " + QPTrim$(TownRec.AppTownOf) + ":" 'Town of Elloree:"
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "         Owners Name:______________________________________________"
    Print #RptHandle, Tab(5); "Business Description:______________________________________________"
    Print #RptHandle, Tab(5); "      Business Phone:______________________________________________"
    Print #RptHandle, Tab(5); "   Federal ID Number:______________________________________________"
    Print #RptHandle, Tab(5); "     State ID Number:______________________________________________"
    Print #RptHandle, Tab(5); "___________________________________________________________________"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "To calculate your " + QPTrim$(TownRec.AppTownOf) + " Business License Fee, Use the"
    Print #RptHandle, Tab(5); "formula below."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "1.  Gross Sales"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "2.  Less Base Amount"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "3.  Excess Gross"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "4.  Base Rate Fee"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "5.  If No. 3 is Greater than"
    Print #RptHandle, Tab(5); "    Zero, divide No. 3 by 1,000"
    Print #RptHandle, Tab(5); "    and round UP"; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle, Tab(5); "6.  Multiply #5 by "; MultiBY$; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle, Tab(5); "7.  Total License Fee # 4 + # 6"; Tab(40); Dash$
    Print #RptHandle, Tab(5); "8.  Add  penalty (" + QPTrim$(Using("$##0.00", TownRec.AppColFee)) + " Collector's"
    If PenAmt = False Then
      Print #RptHandle, Tab(5); "    Fee and " + QPTrim$(Using("#0.00%", TownRec.AppGrsPct / 100)) + " per month after"
    Else
      Print #RptHandle, Tab(5); "    Fee and " + QPTrim$(Using("$##,##0.00", TownRec.AppGrsPct)) + " per month after"
    End If
    Print #RptHandle, Tab(9); QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay); Tab(40); Dash$
    Print #RptHandle, Tab(5); "9.  TOTAL DUE (# 7 + # 8)"; Tab(40); Dash$
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(5); "This is to certify that the amount of total gross for the business"
    Print #RptHandle, Tab(5); "transacted at or through the above location for the calendar year"
    Print #RptHandle, Tab(5); "ending " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", or the last complete fiscal year is true and"
    Print #RptHandle, Tab(5); "correct, and that this report corresponds with the amount that was"
    Print #RptHandle, Tab(5); "reported to the SC Tax Commission or Insurance Commission and with"
    Print #RptHandle, Tab(5); "the Internal Revenue Service."
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); ""
    Print #RptHandle, Tab(5); Dash$; Tab(40); Dash$
    Print #RptHandle, Tab(5); "Firm Name/ Individual Signature"; Tab(40); "By:"
    Print #RptHandle, Chr$(12);
    AppCnt = AppCnt + 1
    TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1
NotNow6:
  Next cnt
  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #6 processed in text format.")
  
  Exit Sub

'-----------------------------------------------------
PrintCustom7:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))

  tab2 = TownLen / 2
  tab2 = Abs(38 - tab2)
  Tab3 = AddLen / 2
  Tab3 = Abs(38 - Tab3)
  Tab4 = CityLen / 2
  Tab4 = Abs(38 - Tab4)
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow7
    Else
      If CustRec.VALID > XDate Then GoTo NotNow7
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm7
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm7
      End If
    End If
    GoTo NotNow7
PrintForm7:
    Print #RptHandle, ""
    Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(tab2); QPTrim$(TownRec.TownName) '"TOWN OF STEPHENS CITY"
    Print #RptHandle, Tab(24); "BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(31); "For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Please print or type:"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Applicant Name: "; CustRec.BillName; Tab(58); "Phone:"
    Print #RptHandle, Tab(5); "               ----------------------------------------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Trade Name: "; Tab(54); "FEIN or SS#"
    Print #RptHandle, Tab(5); "           --------------------------------------------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Mailing Address:                            Physical Address:"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "------------------------------------------  -----------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Phone:                                      Phone:"
    Print #RptHandle, Tab(5); "      ------------------------------------        -----------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Nature Of Business:"
    Print #RptHandle, Tab(5); "                   ------------------------------------------------------"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Gross receipts                 Estimated                         Actual"
    Print #RptHandle, Tab(5); "for year ending"
    Print #RptHandle, Tab(5); QPTrim$(TownRec.AppFiscMonth) + " " + CStr(TownRec.AppFiscDay) + ", " + QPTrim$(YrUpDown(2)); Tab(28); "       -----------                     -----------"
    Print #RptHandle, Tab(5); "(Wholesalers Only...Enter Purchases)"
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "CONTRACTORS ONLY"
    Print #RptHandle, Tab(5); "Please Note: All contractors must have valid Workmans Compensation coverage"
    Print #RptHandle, Tab(5); "in effect for the time period covered by this license. Failure to have"
    Print #RptHandle, Tab(5); "proper coverage will cause your license to be revoked."
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "____ I certify that I am in compliance with the provisions of the Virginia"
    Print #RptHandle, Tab(5); "Workmans Compensation Act, and I will notify the " + QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(5); "if this coverage lapses during the period that this license is in effect."
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "I hereby swear (or affirm) that the statements are true, full and correct to"
    Print #RptHandle, Tab(5); "the best of my knowledge."
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "___________________________________________              ________________"
    Print #RptHandle, "                    Signature                                      Date      "
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "*************************************************************************"
    Print #RptHandle, Tab(5); "FOR OFFICE USE ONLY"
    Print #RptHandle, Tab(5); "Zoning classification approved for this type of business"
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "Approved by ______________________________              ________________"
    Print #RptHandle, "                          Signature                               Date      "

    Print #RptHandle, Chr$(12);
    AppCnt = AppCnt + 1
    TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1

NotNow7:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #7 processed in text format.")
  
  Exit Sub
  
PrintCustom8:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))
  AddLen = Len(QPTrim$(TownRec.AppAdd1))
  CityLen = Len(QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip))
  
  ThisTab = TownLen / 2
  ThisTab = Abs(41 - ThisTab)
  tab2 = Len(QPTrim$(TownRec.AppMayorCouncil))
  tab2 = tab2 / 2
  tab2 = Abs(41 - tab2)
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow8
    Else
      If CustRec.VALID > XDate Then GoTo NotNow8
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm8
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm8
      End If
    End If
    GoTo NotNow8
PrintForm8:
      Print #RptHandle, ""
      Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(ThisTab); QPTrim$(TownRec.AppTownOf)  '"CITY OF ATMORE"
      Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppMayorCouncil)
      Print #RptHandle, ""
      Print #RptHandle, Tab(60); "Date: "; Date$
      Print #RptHandle, ""
      Print #RptHandle, Tab(2); "NOTICE FOR RENEWAL OF BUSINESS LICENSE FOR PERIOD ENDING: " + QPTrim$(TownRec.AppFiscMonth) + ", " + QPTrim$(YrUpDown(1))
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "Business Account # "; IdxRecs(cnt)
      Print #RptHandle, Tab(5); CustRec.BillName
      Print #RptHandle, Tab(5); CustRec.ADDRESS1
      Print #RptHandle, Tab(5); CustRec.ADDRESS2
      Print #RptHandle, Tab(5); RTrim$(CustRec.City); " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode)
      Print #RptHandle, ""
      Print #RptHandle, String$(79, "-")
      Print #RptHandle, Tab(2); "Code"; Tab(9); "Type of License"
      Print #RptHandle, String$(79, "-")
      Lp = 17
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT1)) = 0 Then GoTo Next2
      CatCode$ = QPTrim$(CustRec.BILLCAT1)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT1;
      Print #RptHandle, Tab(9); CustRec.DESC1; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
Next2:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo Next3 ' EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT2)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT2;
      Print #RptHandle, Tab(9); CustRec.DESC2; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
Next3:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo Next4 'EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT3)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT3;
      Print #RptHandle, Tab(9); CustRec.DESC3; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
Next4:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo Next5 'EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT4)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT4;
      Print #RptHandle, Tab(9); CustRec.DESC4; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        Lp = Lp + 1
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
      If Lp >= 54 Then 'if this customer has 4 full categories (10 lines each) then
      'if the page does not break here it will run over with the fifth code
        GoSub PrintHeader8
      End If
Next5:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo EndAtmore1
      CatCode$ = QPTrim$(CustRec.BILLCAT5)
      GoSub GetCode
      Print #RptHandle, Tab(2); CustRec.BILLCAT5;
      Print #RptHandle, Tab(9); CustRec.DESC5; Tab(55); "BASIS AMT"; Tab(69); "LICENSE AMT"
      Lp = Lp + 1
      If CodeType$ = "S" Then
        Print #RptHandle, Tab(2); "Min Due"; Tab(11); "For Recpts Up To"; Tab(31); "Plus"; Tab(37); "Of Recpts Over"
        If BaseAmt1# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$##,###,##0.00", Maximum1#)
          Lp = Lp + 1
        End If
        If BaseAmt2# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$##,###,##0.00", Maximum2#)
          Lp = Lp + 1
        End If
        If BaseAmt3# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$##,###,##0.00", Maximum3#)
          Lp = Lp + 1
        End If
        If BaseAmt4# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$##,###,##0.00", Maximum4#)
          Lp = Lp + 1
        End If
        If BaseAmt5# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$##,###,##0.00", Maximum5#)
          Lp = Lp + 1
        End If
        If BaseAmt6# > 0 Then
          Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$##,###,##0.00", Maximum6#)
          Lp = Lp + 1
        End If
        Print #RptHandle, ; Tab(54); "___________ "; Tab(68); "____________ "
        Lp = Lp + 1
      End If
      If CodeType$ = "F" Then
        Print #RptHandle, Tab(55); "Flat Fee: "; Tab(66); Using("$#,###,##0.00", Amt#)
        Lp = Lp + 1
      End If
      If CodeType$ = "M" Then
        Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
        Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(55); "***N/A***"; Tab(67); "_____________"
        Lp = Lp + 2
      End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
EndAtmore1:
      If Lp >= 36 Then
        GoSub PrintHeader8
      End If

      Print #RptHandle,
      Print #RptHandle, Tab(5); "Make Checks Payable To:"; Tab(45); "License Total: _________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppTownOf); Tab(45); "Penalty:       _________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppAdd1); Tab(45); "Interest:      _________________"
      If QPTrim$(TownRec.SpareSpace) = "" Then
        Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity); Tab(45); "Issue Fee: "; Tab(67); Using("$##0.00", TownRec.IssFee) + " " + QPTrim$(TownRec.SpareSpace)
      Else
        Print #RptHandle, Tab(5); QPTrim$(TownRec.AppCity); Tab(45); "Issue Fee:    " + Using("$##0.00", TownRec.IssFee) + " " + QPTrim$(TownRec.SpareSpace)
      End If
      Print #RptHandle, Tab(5); QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip); Tab(45); "               -----------------"
      Print #RptHandle, Tab(5); ""; Tab(45); "Total Due:     _________________"
      Print #RptHandle,
      Print #RptHandle, Tab(5); "License renewals are due " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + " and delinquent after ";
      Print #RptHandle, QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay)
      
      If PenAmt = False Then
        If TownRec.AppGrsPct = 8 Or TownRec.AppGrsPct = 11 Then
          Print #RptHandle, Tab(5); "at which time an ";
        Else
          Print #RptHandle, Tab(5); "at which time a ";
        End If
        
        Print #RptHandle, CStr(TownRec.AppGrsPct) + "% penalty will be charged. Renewals after " + QPTrim(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay)
        
        If TownRec.AppDiscPct = 8 Or TownRec.AppDiscPct = 11 Then
          Print #RptHandle, Tab(5); "will be charged an " + CStr(TownRec.AppDiscPct) + "% penalty. If you have any questions regarding this"
        Else
          Print #RptHandle, Tab(5); "will be charged a " + CStr(TownRec.AppDiscPct) + "% penalty. If you have any questions regarding this"
        End If
      Else 'amount
        Print #RptHandle, Tab(5); "at which time a ";
        Print #RptHandle, QPTrim$(Using$("$##,##0.00", TownRec.AppGrsPct)) + " penalty will be charged. Renewals after " + QPTrim(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay)
        Print #RptHandle, Tab(5); "will be charged a " + QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + " penalty. If you have any questions regarding this"
      End If
      
      Print #RptHandle, Tab(5); "notice, please call " + QPTrim$(TownRec.AppPhone) + "."
      Print #RptHandle,
      Print #RptHandle, Tab(10); "RENEWALS THAT DO NOT CONTAIN SIGNATURE AND GROSS RECEIPTS"
      Print #RptHandle, Tab(10); "(WHERE REQUIRED) WILL NOT BE PROCESSED."
      Print #RptHandle,
      Print #RptHandle, Tab(5); "I CERTIFY THAT THE ABOVE INFORMATION IS CORRECT"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "NAME ________________________________ TITLE ________________________"
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "SUBSCRIBED AND SWORN TO BEFORE ME THIS ______ DAY OF ______, ______."
      Print #RptHandle, ""
      Print #RptHandle, Tab(5); "NOTARY PUBLIC ____________________________________________"
      Print #RptHandle,
      Print #RptHandle,

      Print #RptHandle, Chr$(12);
      AppCnt = AppCnt + 1
      TempCustRec.CustRecNum = IdxRecs(cnt) 'savein case a reprint is needed
      TempCustRec.AppType = AppType
      TempCustRec.ThisYear = Year$
      TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
      For x = 1 To 5
        TempCustRec.Fee(x) = 0
        TempCustRec.CatCode(x) = ""
        TempCustRec.CatDesc(x) = ""
      Next x
      TempCustRec.MiscNum = Date2Num(Date$)
      TempCustRec.IssFee = IssFee#
      Put TempHandle, Nextcnt, TempCustRec
      Nextcnt = Nextcnt + 1
NotNow8:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #8 processed in text format.")
  
  Exit Sub

PrintHeader8:
  Print #RptHandle, Chr$(12)
  Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.AppTownOf)  '"CITY OF ATMORE"
  Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppMayorCouncil)
  Print #RptHandle,
  Lp = 3

  Return
'-----------------------------------------------------
PrintCustom9:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  TownLen = Len(QPTrim$(TownRec.AppTownOf))

  tab2 = TownLen / 2
  tab2 = Abs(42 - tab2)
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow9
    Else
      If CustRec.VALID > XDate Then GoTo NotNow9
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm9
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm9
      End If
    End If
    GoTo NotNow9
PrintForm9:
    Print #RptHandle, ""
    Print #RptHandle, Tab(2); CStr(Nextcnt)
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(tab2); QPTrim$(TownRec.AppTownOf)
    Print #RptHandle, Tab(26); "   BUSINESS LICENSE APPLICATION"
    Print #RptHandle, Tab(26); "         For Year: "; QPTrim$(YrUpDown(1))
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(3); "Business Name: "; QPTrim$(CustRec.CustName)
    Print #RptHandle, Tab(3); "              -------------------------------------------------------------"
    Print #RptHandle, Tab(3); "Mailing Address: "; QPTrim$(CustRec.ADDRESS1)
    Print #RptHandle, Tab(3); "                 "; QPTrim$(CustRec.City); " "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
    Print #RptHandle, Tab(3); "                -----------------------------------------------------------"
    Print #RptHandle, Tab(3); "Business Address: "
    Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
    Print #RptHandle, Tab(3); "Telephone Number: "
    Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
    Print #RptHandle, Tab(3); "Type of Business: "
    Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
    Print #RptHandle, Tab(3); "Social Security Number:"
    Print #RptHandle, Tab(3); "                       ----------------------------------------------------"
    Print #RptHandle, Tab(3); "Federal Identification Number: "
    Print #RptHandle, Tab(3); "                              ---------------------------------------------"
    Print #RptHandle, Tab(3); "Gross Income Previous Year:"
    Print #RptHandle, Tab(3); "                           ------------------------------------------------"
    Print #RptHandle, Tab(3); "License as Calculated:"
    Print #RptHandle, Tab(3); "                      -----------------------------------------------------"
    If PenAmt = False Then
      Print #RptHandle, Tab(3); QPTrim$(Using("##0", TownRec.AppDiscPct)) + "% Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"
      Print #RptHandle, Tab(3); "                                   ----------------------------------------"
      Print #RptHandle, Tab(3); QPTrim$(Using("##0", TownRec.AppPct)) + "% Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"
    Else
      Print #RptHandle, Tab(3); QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + " Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"
      Print #RptHandle, Tab(3); "                                   ----------------------------------------"
      Print #RptHandle, Tab(3); QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + " Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"
    End If
    Print #RptHandle, Tab(3); "                                 ------------------------------------------"
    Print #RptHandle, Tab(3); "TOTAL AMOUNT DUE: "
    Print #RptHandle, Tab(3); "                 ----------------------------------------------------------"
    Print #RptHandle, Tab(3); ""
    Print #RptHandle, Tab(3); ""
    Print #RptHandle, Tab(3); "   This is to certify that the above is a true statement of the business"
    Print #RptHandle, Tab(3); "transacted at or through the above location for the calendar year ending"
    Print #RptHandle, Tab(3); QPTrim$(TownRec.AppFiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppFiscDay)) + ", " + QPTrim$(YrUpDown(2)); ", and that the report corresponds with the records with"
    Print #RptHandle, Tab(3); "the S.C. Tax Commission of Insurance Commissioner and with the Collector of"
    Print #RptHandle, Tab(3); "Internal Revenue of the United States. I understand that the Town Ordinance"
    Print #RptHandle, Tab(3); "provides for penalties of making false or fraudulent statements in this"
    Print #RptHandle, Tab(3); "application. All licenses are subject to being audited. Failure to provide"
    Print #RptHandle, Tab(3); "all information requested will result in an audit from all required sources."
    Print #RptHandle, Tab(3); ""
    Print #RptHandle, Tab(3); "___________________________________________________________________________"
    Print #RptHandle, Tab(3); "Signature                       Title                              Date"
    Print #RptHandle, Tab(3); ""
    Print #RptHandle, Tab(3); ""
    Print #RptHandle, Tab(3); "FOR OFFICE USE ONLY                        PLEASE REMIT TO:"
    Print #RptHandle, Tab(3); "SIC CODE___________________ "; Tab(46); QPTrim$(TownRec.AppTownOf) 'TOWN OF HEMINGWAY"
    Print #RptHandle, Tab(3); "RATE CLASS_________________ "; Tab(46); QPTrim$(TownRec.AppAdd1)               'P.O. BOX 968"
    Print #RptHandle, Tab(3); "LICENSE NUMBER_____________ "; Tab(46); QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip) 'HEMINGWAY S.C. 29554"
    Print #RptHandle, Chr$(12);
    AppCnt = AppCnt + 1
    TempCustRec.CustRecNum = IdxRecs(cnt) 'save in case a reprint is needed
    TempCustRec.AppType = AppType
    TempCustRec.ThisYear = Year$
    TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
    For x = 1 To 5
      TempCustRec.Fee(x) = 0
      TempCustRec.CatCode(x) = ""
      TempCustRec.CatDesc(x) = ""
    Next x
    TempCustRec.MiscNum = 0
    TempCustRec.IssFee = IssFee#
    Put TempHandle, Nextcnt, TempCustRec
    Nextcnt = Nextcnt + 1
NotNow9:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("Application #9 processed in text format.")
  
  Exit Sub

'-----------------------------------------------------
PrintStandard:
'  ReportFile$ = "CUSTAPPS.RPT"
'  RptHandle = FreeFile
'
'  Open ReportFile$ For Output As #RptHandle
  CodeHandle = FreeFile
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  AppCnt = 0
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  For cnt = 1 To NumOfIdxRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow10
    Else
      If CustRec.VALID > XDate Then GoTo NotNow10
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintSTDForm
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintSTDForm
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfIdxRecs 'NumOfCustRecs
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
NotNow10:
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True

  Close         'Close all open files now
  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile$, "Applications", True
  End If
  
  KillFile ReportFile$
  MainLog ("Standard application processed in text format.")
  Exit Sub

PrintSTDForm:
  GoSub GetCustFee
  AppCnt = AppCnt + 1
  TempCustRec.CustRecNum = IdxRecs(cnt)
  TempCustRec.AppType = AppType
  TempCustRec.ThisYear = Year$
  TempCustRec.AmtPct = QPTrim$(TownRec.UseAmtPctYN)
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  For ll = 1 To 5
    Print #RptHandle, ""
  Next
  Print #RptHandle, 'Tab(37 - tab1); 'Heading1$
  Print #RptHandle, Tab(2); CStr(Nextcnt) 'Tab(37 - tab2); 'Heading2$
  Print #RptHandle, 'Tab(37 - Tab3); 'Heading3$
  Print #RptHandle, 'Tab(37 - Tab4); 'Heading4$
  Print #RptHandle, 'Tab(66); Year$ ' Form$(2, 0)
  Print #RptHandle,
  Print #RptHandle, Tab(11); QPTrim$(CustRec.BillName)
  Print #RptHandle, Tab(11); QPTrim$(CustRec.ADDRESS1)
  Print #RptHandle, Tab(11); QPTrim$(CustRec.ADDRESS2)
  Print #RptHandle, Tab(11); RTrim$(CustRec.City); "  "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(11); QPTrim$(CustRec.CustName)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT1);
  Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1);
  Print #RptHandle, Tab(62); Using("##,##0.00", FeeAmt1#)
  TempCustRec.CatCode(1) = QPTrim$(CustRec.BILLCAT1)
  TempCustRec.CatDesc(1) = QPTrim$(CustRec.DESC1)
  TempCustRec.Fee(1) = FeeAmt1#
  SCnt = 24
  If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT2);
  Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2);
  Print #RptHandle, Tab(62); Using("##,##0.00", FeeAmt2#)
  TempCustRec.CatCode(2) = QPTrim$(CustRec.BILLCAT2)
  TempCustRec.CatDesc(2) = QPTrim$(CustRec.DESC2)
  TempCustRec.Fee(2) = FeeAmt2#
  SCnt = 25
  If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT3);
  Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3);
  Print #RptHandle, Tab(62); Using("##,##0.00", FeeAmt3#)
  TempCustRec.CatCode(3) = QPTrim$(CustRec.BILLCAT3)
  TempCustRec.CatDesc(3) = QPTrim$(CustRec.DESC3)
  TempCustRec.Fee(3) = FeeAmt2#
  SCnt = 26
  If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT4);
  Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4);
  Print #RptHandle, Tab(62); Using("##,##0.00", FeeAmt4#)
  TempCustRec.CatCode(4) = QPTrim$(CustRec.BILLCAT4)
  TempCustRec.CatDesc(4) = QPTrim$(CustRec.DESC4)
  TempCustRec.Fee(4) = FeeAmt2#
  SCnt = 27
  If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo ExitFormPrint
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT5);
  Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5);
  Print #RptHandle, Tab(62); Using("##,##0.00", FeeAmt5#)
  TempCustRec.CatCode(5) = QPTrim$(CustRec.BILLCAT5)
  TempCustRec.CatDesc(5) = QPTrim$(CustRec.DESC5)
  TempCustRec.Fee(5) = FeeAmt5#
  SCnt = 28
  
ExitFormPrint:
  If IssFee > 0 Then
    Print #RptHandle, Tab(15); "Issuance Fee"; Tab(62); Using$("##,##0.00", IssFee#)
  End If
  TotalFees = OldRound(FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + IssFee#)
  TempCustRec.MiscNum = TotalFees
  TempCustRec.IssFee = IssFee#

  For LCnt = SCnt To 35
    Print #RptHandle, ""
  Next
  Print #RptHandle, Tab(62); Using("##,##0.00", TotalFees)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(62); Using("##,##0.00", TotalFees)
  Print #RptHandle,
  Print #RptHandle,
  TotalCust = TotalCust + 1
  Put TempHandle, Nextcnt, TempCustRec
  Nextcnt = Nextcnt + 1

Return

GetCode:
  For Snt& = 1 To NumOfARCatRecs
    Get CodeHandle, Snt&, CodeRec
    If QPTrim$(CodeRec.CatCode) = CatCode$ Then
      CODEDESC$ = QPTrim$(CodeRec.CODEDESC)
      Select Case CodeRec.CodeType
      Case "F"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case "M"
        DESC1$ = "Per Each"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case Is = "S"
        BaseAmt1# = CodeRec.BaseAmt1
        Revenue1# = CodeRec.Recpt1
        Percent1# = CodeRec.Percent1
        Maximum1# = CodeRec.Maximum1
        BaseAmt2# = CodeRec.BaseAmt2
        Revenue2# = CodeRec.Recpt2
        Percent2# = CodeRec.Percent2
        Maximum2# = CodeRec.Maximum2
        BaseAmt3# = CodeRec.BaseAmt3
        Revenue3# = CodeRec.Recpt3
        Percent3# = CodeRec.Percent3
        Maximum3# = CodeRec.Maximum3
        BaseAmt4# = CodeRec.BaseAmt4
        Revenue4# = CodeRec.Recpt4
        Percent4# = CodeRec.Percent4
        Maximum4# = CodeRec.Maximum4
        BaseAmt5# = CodeRec.BaseAmt5
        Revenue5# = CodeRec.Recpt5
        Percent5# = CodeRec.Percent5
        Maximum5# = CodeRec.Maximum5
        BaseAmt6# = CodeRec.BaseAmt6
        Revenue6# = CodeRec.Recpt6
        Percent6# = CodeRec.Percent6
        Maximum6# = CodeRec.Maximum6
        CodeType$ = CodeRec.CodeType
      Case Else
        CodeType$ = "N"
      End Select
      Exit For
    End If
  Next Snt&


GotCode:
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
  CustFee# = OldRound#(CustFee# + FeeAmt3#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt4#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
'  FeeAmt# = 0
  
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintAppsRenwls", "PrintText", Erl)
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
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim Code$, ll As Integer
  Dim Year$
  Dim AppFormat$
  Dim ReturnAdd$
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType 'CustSearchNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim ThisCode$, SCnt As Integer
  Dim TotalCust As Integer
  Dim ReportFile$, RptHandle As Integer
  Dim LicTotal#
  Dim CatCode$
  Dim ZZCnt As Integer
  Dim Snt&, Amt#
  Dim CODEDESC$
  Dim CodeType$
  Dim DESC1$
  Dim BaseAmt1#, BaseAmt2#, BaseAmt3#, BaseAmt4#, BaseAmt5#, BaseAmt6#
  Dim Revenue1#, Revenue2#, Revenue3#, Revenue4#, Revenue5#, Revenue6#
  Dim Percent1#, Percent2#, Percent3#, Percent4#, Percent5#, Percent6#
  Dim Maximum1#, Maximum2#, Maximum3#, Maximum4#, Maximum5#, Maximum6#
  Dim Nextcnt As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim AppType As Integer
  Dim BaseFee$
  Dim TotalFees As Double
  Dim AppCnt As Integer
  Dim MultiBY$
  Dim dlm$
  Dim NumOfCats As Integer
  Dim YrUpDown$(1 To 10)
  Dim CustFee#
  Dim FeeAmt1#, FeeAmt2#, FeeAmt3#, FeeAmt4#, FeeAmt5#
  Dim Prorate#
  Dim Mult#
  Dim Revenue#
  Dim IssFee#
  Dim XDate As Integer
  Dim RangeFlag As Integer
  Dim Laser5 As LaserLetterType5
  Dim LHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  RangeFlag = 2
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  End If
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  XDate = Date2Num(fptxtNewXDate)
  
  For x = 1 To 10
    YrUpDown$(x) = "0000"
  Next x
  
  dlm = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  IssFee# = TownRec.IssFee
  AppType = TownRec.AppForm
  
  Nextcnt = 1
'  KillFile BLTempCustRecName 'kill temporary file used for reprints

  Code$ = QPTrim$(fptxtCatCode.Text)
  Year$ = fptxtLicYear.Text

'  OpenSrchNameIdxFile IdxHandle
  OpenCustNameIdxFile IdxHandle

  NumOfIdxRecs = LOF(IdxHandle) / Len(CustIdx)
  If NumOfIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim IdxRecs(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, CustIdx
    IdxRecs(x) = CustIdx.CustRec
  Next x
  Close IdxHandle

  OpenCustFile CHandle

'  ReportFile$ = "CUSTAPPS.RPT"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
  
  If AppType > 1 Then
    OpenCatCodeFile CodeHandle
    NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
    If AppType = 2 Then
      GoSub PrintCustom2
    ElseIf AppType = 3 Then
      GoSub PrintCustom3
    ElseIf AppType = 4 Then
      GoSub PrintCustom4
    ElseIf AppType = 5 Then
      GoSub PrintCustom5
    ElseIf AppType = 6 Then
      GoSub PrintCustom6
    ElseIf AppType = 7 Then
      GoSub PrintCustom7
    ElseIf AppType = 8 Then
      GoSub PrintCustom8
    ElseIf AppType = 9 Then
      GoSub PrintCustom9
    ElseIf AppType = 10 Then
      GoSub PrintCustom10
    End If
  Else
    GoSub PrintStandard
  End If


PrintCustom2: 'all subs close out of this form
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP2.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow2
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow2
    Else
      If CustRec.VALID > XDate Then GoTo NotNow2
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm2
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm2
      End If
    End If
    GoTo NotNow2
PrintForm2:
   AppCnt = AppCnt + 1
   '                              0
   Print #RptHandle, "LICENSE   APPLICATION"; dlm;
   '                                        1                                    2
   Print #RptHandle, "     " + QPTrim$(TownRec.AppTownOf); dlm; "ACCOUNT NO.    " + Using("####0", IdxRecs(cnt)); dlm;
   '                      3
   Print #RptHandle, "START DATE:  " + UCase(QPTrim$(TownRec.AppStartMonth)) + "   " + CStr(TownRec.AppStartDay) + ", " + YrUpDown(1); dlm;
   '                                  4
   Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
   '                                  5
   Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm;
   '                                     6
   Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
   '                                     7
   Print #RptHandle, QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + "  " + QPTrim$(CustRec.ZipCode); dlm;
   '                                     8
   Print #RptHandle, "     " + "TAX  MAP________    BLOCK________   LOT_________        ZONING  DISTRICT_____________________"; dlm;
   '                                     9
   Print #RptHandle, "     " + "FEDERAL  ID/SS  NUMBER_____________________     " + QPTrim$(TownRec.AppState) + " TAX  ID  NUMBER_______________________"; dlm;
   '                                     10
   Print #RptHandle, "     " + "TYPE  OF  BUSINESS:_____________________________________________________________________"; dlm;
   '                                     11
   Print #RptHandle, "     " + "APPLICATION  FOR:____  NEW____  RENEWAL____  GOING  OUT  OF  BUSINESS(DATE)____________"; dlm;
   '                                     12
   Print #RptHandle, "     " + "OWNERSHIP:_____  CORPORATION_____  PARTNERSHIP______  INDIVIDUAL-NO EMPLOYEES______"; dlm;
   '                                     13
   Print #RptHandle, "     " + "NAME  OF  OWNER,  PARTNER  OR  PRINCIPAL_______________________________________________"; dlm;
   '                                     14
   Print #RptHandle, "     " + "TELEPHONE  NO.  LOCAL:________________  HOME:_______________  EMERGENCY:______________"; dlm;
   '                                     15
   Print #RptHandle, "     " + "FAX  NO.________________   E-MAIL:________________________________________________________"; dlm;
   '                                     16
   Print #RptHandle, "     " + "IS  HAZARDOUS  WASTE  INVOLVED  IN  OPERATION?  ______NO   ______YES     (ATTACH  DETAILS)"; dlm;
   '                                     17
   Print #RptHandle, "     " + "CODE  CLEARANCE: ___ZONING  ____INSPECTION  ___FIRE  ____HEALTH  ____LAW  ENFORCEMENT"; dlm;
   '                            18
   Print #RptHandle, "COMPUTATION  OF  LICENSE  TAX"; dlm;
   '                                     19
   Print #RptHandle, "     " + "COMPUTE   LICENSE   TAX  ACCORDING  TO  THE   FOLLOWING   SCHEDULE   AND   MAKE   CHECKS"; dlm;
   '                                       20
   Print #RptHandle, "     " + "PAYABLE  TO :  " + QPTrim$(TownRec.AppTownOf) + ".  DELIVER  BY  DUE  DATE :  " + UCase(QPTrim$(TownRec.AppLicRetMonth)) + "  " + CStr(TownRec.AppLicRetDay) + ",  " + YrUpDown(2); dlm;
   '                                       21
   Print #RptHandle, "     " + "GROSS  INCOME  FOR  PRECEDING  CALENDAR  OR  FISCAL  YEAR....................$_________________"; dlm;
   '                                       22
   Print #RptHandle, "     " + "LESS  INCOME  ON  WHICH  A  LICENSE  TAX  WAS  PAID  TO  ANOTHER"; dlm;
   '                                       23
   Print #RptHandle, "     " + "CITY  OR  COUNTY  FOR  OPERATIONS  OUTSIDE  CITY/COUNTY.........................$_________________"; dlm;
   '                                       24
   Print #RptHandle, "     " + "BALANCE  OF  GROSS  INCOME  SUBJECT  TO  LICENSE  TAX..............................$_________________"; dlm;
   '                                       25
   Print #RptHandle, "     " + "TAX:    RATE  CLASS  MINIMUM  ON  FIRST  " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(1))) + ":  " + QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(1))) + "  PLUS"; dlm;
   '                                       26
   Print #RptHandle, "     " + QPTrim$(Using("$#,###,##0.00", TownRec.AppBaseFee(2))) + "  PER  " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(2))) + "  FOR  INCOME  OVER  " + QPTrim$(Using("$#,###,##0.00", TownRec.AppGrsRcpts(3))); dlm;
   '                                       27
   Print #RptHandle, "     " + "[See  declining  rate  schedule  for  over  $1  million]                                                        [OFFICE  USE  ONLY]"; dlm;
   '                                       28
   Print #RptHandle, "                                                                                  TOTAL  LICENSE  TAX $_________  [PAYMENT  RECORD]"; dlm;
   
   If PenAmt = False Then
     '                                       29
     Print #RptHandle, "     " + "PENALTY  AFTER  DUE  DATE IS " + CStr(TownRec.AppPct) + " %  PER  MONTH  $____________      [CHECK  NO.  _______________]"; dlm;
   Else
     '                                       29
     Print #RptHandle, "     " + "PENALTY  AFTER  DUE  DATE IS " + QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + "  PER  MONTH  $____________  [CHECK  NO. ______________]"; dlm;
   End If
   
   '                                       30
   Print #RptHandle, "     " + "TOTAL  LICENSE  TAX  AND  PENALTY  $_____________                       [DATE  RECEIVED____________]"; dlm;
   '                       31
   Print #RptHandle, "CERTIFICATION"; dlm;
   '                                         32
   Print #RptHandle, "     " + "I  (WE)  DO  CERTIFY  THAT  THE  ABOVE  INFORMATION  AND  AMOUNT  RETURNED  AS  GROSS"; dlm;
   '                                         33
   Print #RptHandle, "     " + "INCOME  FROM  MY  BUSINESS  IS  TRUE  AND  CORRECT  AND  I  HAVE  MADE  NO  DEDUCTIONS"; dlm;
   '                                         34
   Print #RptHandle, "     " + "EXCEPT  INCOME  ON  WHICH  I  HAVE  PAID  BUSINESS  LICENSE  TAX  TO  ANOTHER  CITY  OR"; dlm;
   '                                         35
   Print #RptHandle, "     " + "COUNTY,  FOR  WHICH  I  HAVE  PROOF  OF  PAYMENT.  I  AM  FAMILIAR  WITH  THE  PENALTY"; dlm;
   '                                         36
   Print #RptHandle, "     " + "PROVISIONS  OF  THE  ORDINANCE  AND  GROUNDS  FOR  LICENSE  REVOCATION,  INCLUDING"; dlm;
   '                                         37
   Print #RptHandle, "     " + "MAKING  FALSE  OR  FRAUDULENT  STATEMENTS  IN  THIS  APPLICATION.  I  CERTIFY  THAT"; dlm;
   '                                         38
   Print #RptHandle, "     " + "ALL  BUSINESS  PERSONAL  PROPERTY  TAXES  AND  PAYABLES  DUE  TO  THE  CITY / COUNTY"; dlm;
   '                                         39
   Print #RptHandle, "     " + "HAVE  BEEN  PAID,  AND  THAT  THE  ABOVE  BUSINESS  NAME  IS  THE  SAME  AS  REPORTED"; dlm;
   '                                         40
   Print #RptHandle, "     " + "ON  DOCUMENTS  FILED  WITH  THE  STATE  AND  FEDERAL  GOVERNMENTS.  I  UNDERSTAND  MY"; dlm;
   '                                         41
   Print #RptHandle, "     " + "BUSINESS  INCOME  TAX  RETURNS  AND  OTHER  DOCUMENTS  MAY  BE  INSPECTED  TO  VERIFY"; dlm;
   '                                         42
   Print #RptHandle, "     " + "GROSS  INCOME  OR  OTHER  BUSINESS  DATA."; dlm;
   '                                         43
   Print #RptHandle, "     " + "_______________________________________________________________________________________"; dlm;
   '                                         44
   Print #RptHandle, "     " + "SIGNATURE                                                               TITLE                                                           DATE"
   Nextcnt = Nextcnt + 1
NotNow2:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp2.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #2 processed in graphics format.")
  
  Exit Sub
''-----------------------------------------------------

PrintCustom3:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP3.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow3
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow3
    Else
      If CustRec.VALID > XDate Then GoTo NotNow3
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm3
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm3
      End If
    End If
    GoTo NotNow3
PrintForm3:
    '                               0
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                                1
    Print #RptHandle, "BUSINESS  LICENSE  APPLICATION"; dlm;
    '                                2
    Print #RptHandle, "For  Year:  " + QPTrim$(YrUpDown(1)); dlm;
    '                              3                             4
    Print #RptHandle, QPTrim$(CustRec.CustName); dlm; QPTrim$(CustRec.CustNumb); dlm;
    '                             5
    Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
    '                         6                                7
    Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm; QPTrim$(CustRec.City) + ", " + CustRec.State + "  " + QPTrim$(CustRec.ZipCode); dlm;
    '                         8
    If QPTrim$(CustRec.WPHONE) = "(" Then CustRec.WPHONE = ""
    Print #RptHandle, QPTrim$(CustRec.WPHONE); dlm;
    Rem 22 lines printed here
    '                         9
    Print #RptHandle, "TYPE  OF  BUSINESS  LICENSE  APPLYING  FOR:"; dlm;
    
    If TownRec.IssFee > 0 Then
      '                         10
      Print #RptHandle, "_____________ Contracting  or  Construction  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + "  plus  " + QPTrim(Using("$#,##0.00", TownRec.IssFee)) + " Issuance  Fee."; dlm;
      '                         11
      Print #RptHandle, "_____________ Retail  Sales  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + "  plus  " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + "  of  " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + "  of  gross  receipts"; dlm;
      '                         12
      Print #RptHandle, "                           over  " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + "  plus  " + QPTrim$(Using("$#, ##0.00", TownRec.IssFee)) + "  Issuance  Fee."; dlm;
      '                         13
      Print #RptHandle, "_____________ Financial,  Real  Estate  or  Professional  Service  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + "  plus  " + QPTrim(Using("$#,##0.00", TownRec.IssFee)); dlm;
      '                         14
      Print #RptHandle, "                           Issuance  Fee."; dlm;
      '                         15
      Print #RptHandle, "_____________ Repair,  Personal,  Business  or  Delivery  Service  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + "  plus  " + QPTrim(Using("$#,##0.00", TownRec.IssFee)); dlm;
      '                         16
      Print #RptHandle, "                           Issuance  Fee."; dlm;
    Else
      '                         10
      Print #RptHandle, "_____________ Contracting  or  Construction:  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + "."; dlm;
      '                         11
      Print #RptHandle, "_____________ Retail  Sales:  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + "  plus  " + QPTrim$(Using("##0", TownRec.AppNumer)) + "/" + QPTrim$(Using("##0", TownRec.AppDenom)) + "  of  " + QPTrim$(Using("##0%", (TownRec.AppGrsPct / 100))) + "  of  gross  receipts"; dlm;
      '                         12
      Print #RptHandle, "                           over  " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))) + "."; dlm;
      '                         13
      Print #RptHandle, "_____________ Financial,  Real  Estate  or  Professional  Service:  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + "."; dlm;
      '                         14
      Print #RptHandle, "                           "; dlm;
      '                         15
      Print #RptHandle, "_____________ Repair,  Personal,  Business  or  Delivery  Service:  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + "."; dlm;
      '                         16
      Print #RptHandle, "                           "; dlm;
    End If
    '                         17
    Print #RptHandle, "_____________ Other (Specify) ______________________________________________"; dlm;
    '                         18
    Print #RptHandle, "Estimate  of  ______________  gross  receipts  or  preceding  year's  gross "; dlm;
    '                         19
    Print #RptHandle, "receipts  ______________________"; dlm;
    '                         20
    Print #RptHandle, "AMOUNT OF LICENSE TAX FOR  " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ",  THROUGH  " + QPTrim$(TownRec.AppLicRetMonth) + "  " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim$(YrUpDown(2)) + " IS:   $ ________ "; dlm;
    '                         21
    Print #RptHandle, "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "; dlm;
    '                         22
    Print #RptHandle, "ACTIVITY SHALL BE CONDUCTED: "; dlm;
    '                         23
    Print #RptHandle, "I  certify  that  the  statements  and  figures  set  forth  on  this  application"; dlm;
    '                         24
    Print #RptHandle, "are  true  to  the  best  of  my  knowledge."; dlm;
    '                         25
    Print #RptHandle, "Signature  of  Applicant"; dlm;
    If PenAmt = False Then
      '                         26
      Print #RptHandle, "To  Avoid  Late  Penalty  Charge  of  " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + " ,  Renew  Your  License  By  " + QPTrim$(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + ",  " + QPTrim$(YrUpDown(3)) + "."; dlm;
    Else
      '                         26
      Print #RptHandle, "To  Avoid  Late  Penalty  Charge  of  " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + " ,  Renew  Your  License  By  " + QPTrim$(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + ",  " + QPTrim$(YrUpDown(3)) + "."; dlm;
    End If
    '                         27
    Print #RptHandle, "Return  Application  and  Fee  to: "; dlm;
    '                         28
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                         29
    Print #RptHandle, QPTrim$(TownRec.AppAdd1); dlm;
    '                         30
    Print #RptHandle, QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    AppCnt = AppCnt + 1
NotNow3:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp3.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #3 processed in graphics format.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom4:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(4)) = "Curr" Then
    YrUpDown(4) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "+1" Then
    YrUpDown(4) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(4)) = "-1" Then
    YrUpDown(4) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(5)) = "Curr" Then
    YrUpDown(5) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "+1" Then
    YrUpDown(5) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(5)) = "-1" Then
    YrUpDown(5) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(6)) = "Curr" Then
    YrUpDown(6) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "+1" Then
    YrUpDown(6) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(6)) = "-1" Then
    YrUpDown(6) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(7)) = "Curr" Then
    YrUpDown(7) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "+1" Then
    YrUpDown(7) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(7)) = "-1" Then
    YrUpDown(7) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(8)) = "Curr" Then
    YrUpDown(8) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "+1" Then
    YrUpDown(8) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(8)) = "-1" Then
    YrUpDown(8) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(9)) = "Curr" Then
    YrUpDown(9) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "+1" Then
    YrUpDown(9) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(9)) = "-1" Then
    YrUpDown(9) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(10)) = "Curr" Then
    YrUpDown(10) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "+1" Then
    YrUpDown(10) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(10)) = "-1" Then
    YrUpDown(10) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP4.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow4
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow4
    Else
      If CustRec.VALID > XDate Then GoTo NotNow4
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm4
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm4
      End If
    End If
    GoTo NotNow4
PrintForm4:
    '                           0
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                           1
    Print #RptHandle, "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"; dlm; 'line 2
    '                                       2                            3
    Print #RptHandle, "For Year:   "; QPTrim$(YrUpDown(1)); dlm; "PAGE 1"; dlm;
    '                           4
    Print #RptHandle, "Dear  Business  Owner:"; dlm;
    '                           5
    Print #RptHandle, "     For  the  purpose  of  computing  Business,  Professional  and  Occupational"; dlm;
    '                           6
    Print #RptHandle, "License  (BPOL)  Tax  promulgated  by  Virginia  Code  Section  58.1-3700  et  seq."; dlm;
    '                           7
    Print #RptHandle, "and  " + QPTrim$(TownRec.AppCity) + "  Town  Ordinance  #" + QPTrim$(TownRec.AppCityOrd) + "  adopted  "; dlm;
    '                           8
    Print #RptHandle, MakeRegDate(TownRec.AppAdoptDate) + "  please  complete  and  return  this  form  with  the  required"; dlm;
    '                           9
    Print #RptHandle, "information  no  later  than  " + QPTrim$(TownRec.AppDiscMonth) + "  " + CStr(TownRec.AppDiscDay) + ",  " + QPTrim$(YrUpDown(2)) + "."; dlm;
    '                       10
    Print #RptHandle, "Respectfully,"; dlm;
    '                       11
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                                12
    Print #RptHandle, QPTrim$(TownRec.AppMayorCouncil); dlm;
    '                           13
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                            14
    Print #RptHandle, QPTrim$(TownRec.TownAdd1); dlm;
    '                            15
    Print #RptHandle, QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip); dlm;
    '                            16
    Print #RptHandle, "Application  for  Town  Licenses"; dlm; 'line 3
    '                            17
    Print #RptHandle, "For  period  beginning  " + QPTrim$(TownRec.AppStartMonth) + "  " + CStr(TownRec.AppStartDay) + ",  " + QPTrim$(YrUpDown(3)) + "  (or  start  of  business  in  " + QPTrim$(YrUpDown(4)) + ")"; dlm;
    '                            18
    Print #RptHandle, "and  ending  " + QPTrim$(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + ",  " + QPTrim$(YrUpDown(5)); dlm;
    '                          19
    Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
    '                          20
    Print #RptHandle, QPTrim$(CustRec.CustName); dlm;
    '                          21                     22
    Print #RptHandle, "BUSINESS  ADDRESS:"; dlm; "HOME  ADDRESS"; dlm;
    '                              23                               24
    Print #RptHandle, "MAIL:  " + QPTrim$(CustRec.ADDRESS1); dlm; "MAIL: "; dlm;
    '                            25
    Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
    '                                      26
    Print #RptHandle, RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
    '                   27
    Print #RptHandle, "911: "; dlm;
    '                    28
    Print #RptHandle, "PHONE: "; dlm;
    '                                    29
    Print #RptHandle, "A SEPARATE LICENSE WILL BE ISSUED FOR EACH TYPE OF BUSINESS"; dlm;
    '                                    30
    Print #RptHandle, "PERFORMED, AS REQUIRED PER THE " + UCase(QPTrim$(TownRec.AppCityOrd)) + ". THIS WILL NOT"; dlm;
    '                                    31
    Print #RptHandle, "RESULT IN ANY ADDITONAL COST TO BUSINESSES.  PLEASE REPORT GROSS"; dlm;
    '                                    32
    Print #RptHandle, "RECEIPTS FOR EACH CLASSIFICATION THAT APPLIES TO YOUR BUSINESS."; dlm;
    '                           33
    Print #RptHandle, QPTrim$(TownRec.TownName); dlm;
    '                                        34
    Print #RptHandle, "BUSINESS,  PROFESSIONAL  AND  OCCUPATIONAL LICENSE"; dlm;
    '                                        35                           36
    Print #RptHandle, "For  Year:  " + QPTrim$(YrUpDown(1)); dlm; "PAGE  2"; dlm;
    '                         37
    Print #RptHandle, "WHOLESALE  MERCHANT:"; dlm;
    '                         38
    Print #RptHandle, "Gross  Receipts  through  " + CStr(TownRec.AppWholeMonth) + "-" + CStr(TownRec.AppWholeDay) + "-" + QPTrim$(YrUpDown(6)) + " as  shown  by  applicants  records"; dlm;
    '                         39
    Print #RptHandle, "RETAIL  MERCHANT:"; dlm;
    '                         40
    Print #RptHandle, "Gross  Receipts  through " + CStr(TownRec.AppRetailMonth) + "-" + CStr(TownRec.AppRetailDay) + "-" + QPTrim$(YrUpDown(7)) + " as  shown  by  applicants  records"; dlm;
    '                         41
    Print #RptHandle, "FINANCIAL,  REAL  ESTATE  AND  PROFESSIONAL:"; dlm;
    '                         42
    Print #RptHandle, "Gross  Receipts  through " + CStr(TownRec.AppFinMonth) + "-" + CStr(TownRec.AppFinDay) + "-" + QPTrim$(YrUpDown(8)) + " as  shown  by  applicants  records"; dlm;
    '                         43
    Print #RptHandle, "CONTRACTING:"; dlm;
    '                         44
    Print #RptHandle, "Gross  Receipts  through " + CStr(TownRec.AppContMonth) + "-" + CStr(TownRec.AppContDay) + "-" + QPTrim$(YrUpDown(9)) + " as  shown  by  applicants  records"; dlm;
    '                         45
    Print #RptHandle, "(Subject  to  Virginia  Code  Sec  58.1-3715)"; dlm;
    '                         46
    Print #RptHandle, "REPAIR,  PERSONAL  or  BUSINESS  SERVICES:"; dlm;
    '                         47
    Print #RptHandle, "Gross  Receipts  through " + CStr(TownRec.AppRepairMonth) + "-" + CStr(TownRec.AppRepairDay) + "-" + QPTrim$(YrUpDown(10)) + " as  shown  by  applicants  records"; dlm;
    '                         48
    Print #RptHandle, "If  uncertain  of  your  business  classification(s),  please  call  the  Town  Office  at"; dlm;
    '                         49
    Print #RptHandle, QPTrim$(TownRec.AppPhone) + "  for  assistance."; dlm;
    '                         50
    Print #RptHandle, "I  do  affirm  that  the  foregoing  figures  are  true,  complete  and accurate  to  the"; dlm;
    '                         51
    Print #RptHandle, "best of  my  knowledge."; dlm;
    '                         52
    Print #RptHandle, "Signature"; dlm;
    '                       53
    Print #RptHandle, "Print Name"; dlm;
    '                          54
    Print #RptHandle, "*** IMPORTANT ***"; dlm;
    '                            55
    Print #RptHandle, "APPLICATION  MUST  BE  RETURNED  PRIOR  TO " + UCase(QPTrim$(TownRec.AppFiscMonth)) + " " + CStr(TownRec.AppFiscDay) + " OF  EACH  YEAR"; dlm;
    '                            56
    Print #RptHandle, "TO  AVOID  PENALTY. LICENSE  FEES  ARE  DUE  PRIOR  TO  " + UCase(QPTrim$(TownRec.AppLicRetMonth)) + " " + CStr(TownRec.AppLicRetDay); dlm;
    '                            57
    Print #RptHandle, "OF  EACH  YEAR  TO  AVOID  PENALTY  AND  INTEREST. INTENTIONALLY  PROVIDING"; dlm;
    '                            58
    Print #RptHandle, "INSUFFICIENT  OR  INACCURATE  INFORMATION  MAY  RESULT  IN  LEGAL  RECOURSE"; dlm;
    '                            59
    Print #RptHandle, "BY  THE  TOWN  OF " + UCase(QPTrim$(TownRec.AppCity)) + " AS  SET  FORTH  BY  VIRGINIA  CODE."
    AppCnt = AppCnt + 1
NotNow4:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp4.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #4 processed in graphics format.")
  
  Exit Sub
'-----------------------------------------------------

PrintCustom5:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(3)) = "Curr" Then
    YrUpDown(3) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "+1" Then
    YrUpDown(3) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(3)) = "-1" Then
    YrUpDown(3) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP5.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow5
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow5
    Else
      If CustRec.VALID > XDate Then GoTo NotNow5
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm5
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm5
      End If
    End If
    GoTo NotNow5
PrintForm5:
    '                             0
    Print #RptHandle, QPTrim$(TownRec.TownName); dlm;
    '                             1
    Print #RptHandle, "BUSINESS LICENSE APPLICATION"; dlm;
    '                           2
    Print #RptHandle, "For Year: "; QPTrim$(YrUpDown(1)); dlm;
    '                           3
    Print #RptHandle, "Business Name:  " + QPTrim$(CustRec.CustName); dlm;
    '                           4
    Print #RptHandle, "Street Address of Business:  "; dlm;
    '                           5
    Print #RptHandle, "Zoning of Business Location:  "; dlm;
    '                           6
    Print #RptHandle, "Telephone Number:  "; dlm;
    '                           7
    Print #RptHandle, "Applicant's Name:  " + QPTrim$(CustRec.BillName); dlm;
    '                           8
    Print #RptHandle, "Applicant's Address:  " + QPTrim$(CustRec.ADDRESS1); dlm;
    '                           9
    Print #RptHandle, "Telephone Number:  " + QPTrim$(CustRec.WPHONE); dlm;
    '                           10
    Print #RptHandle, "TYPE OF BUSINESS LICENSE APPLYING FOR:"; dlm;
    '                           11
    Print #RptHandle, "_______ Contracting  or  Construction  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(1))) + "  or  " + Using("#.###", TownRec.AppCentsPer(1)) + "  cents  per  " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(1))); dlm;
    '                           12
    Print #RptHandle, "gross  receipts  whichever  is  greater."; dlm;
    '                           13
    Print #RptHandle, "_______ Retail  Sales  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(2))) + "  or  " + Using("#.###", TownRec.AppCentsPer(2)) + "  cents  per  " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(2))) + "  whichever"; dlm;
    '                           14
    Print #RptHandle, "is  greater."; dlm;
    '                           15
    Print #RptHandle, "_______ Financial,  Real  Estate  or  Professional  Service  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(3))) + "  or  "; dlm;
    '                           16
    Print #RptHandle, Using("#.###", TownRec.AppCentsPer(3)) + "  cents  per  " + QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(3))) + "  whichever  is  greater."; dlm;
    '                           17
    Print #RptHandle, "_______ Repair,  Personal  or  Business  Service  " + QPTrim$(Using("$#,##0.00", TownRec.AppBaseFee(4))) + "  or  " + Using("#.###", TownRec.AppCentsPer(4)) + "  cents  per  "; dlm;
    '                           18
    Print #RptHandle, QPTrim$(Using("$##,###,##0.00", TownRec.AppGrsRcpts(4))) + "  whichever  is  greater."; dlm;
    '                           19
    Print #RptHandle, "_______ Other (Specify) "; dlm;
    '                           20
    Print #RptHandle, "Estimate  of  ______________  gross  receipts  or  preceding  year's  gross "; dlm;
    '                           21
    Print #RptHandle, "receipts  ______________________.  Enclose  copy  of  most  recent  schedule  C"; dlm;
    '                           22
    Print #RptHandle, "or  other  comparable  federal  document."; dlm;
    '                           23
    Print #RptHandle, "AMOUNT OF LICENSE TAX FOR " + QPTrim$(TownRec.AppStartMonth) + " " + CStr(TownRec.AppStartDay) + ", THROUGH " + QPTrim$(TownRec.AppLicRetMonth) + " " + CStr(TownRec.AppLicRetDay) + ", " + QPTrim(YrUpDown(2)) + " IS:  $_______"; dlm;
    '                           24
    Print #RptHandle, "ANY SPECIAL CONDITIONS OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED "; dlm;
    '                           25
    Print #RptHandle, "ACTIVITY SHALL BE CONDUCTED:"; dlm;
    '                           26
    Print #RptHandle, "I  certify  that  the  statements  and  figures  set  forth  on  this  application"; dlm;
    '                           27
    Print #RptHandle, "are  true  to  the  best  of  my  knowledge."; dlm;
    '                           28
    Print #RptHandle, "Signature  of  Applicant"; dlm;
    If PenAmt = False Then
      '                           29
      Print #RptHandle, "To  Avoid  Late  Penalty  Charge  of  " + QPTrim(Using("##0%", (TownRec.AppPct / 100))) + ",  Renew  Your  License  By  " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."; dlm;
    Else
      '                           29
      Print #RptHandle, "To Avoid Late Penalty Charge of  " + QPTrim(Using("$##,##0.00", TownRec.AppPct)) + ", Renew Your License By " + QPTrim$(TownRec.AppPenMonth) + " " + CStr(TownRec.AppPenDay) + ", " + QPTrim$(YrUpDown(3)) + "."; dlm;
    End If
    '                           30
    Print #RptHandle, "Return Application and Fee to:"; dlm;
    '                           31
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                           32
    Print #RptHandle, QPTrim$(TownRec.AppAdd1); dlm;
    '                           33
    Print #RptHandle, QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + "  " + QPTrim$(TownRec.AppZip)
    AppCnt = AppCnt + 1
NotNow5:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp5.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #5 processed in graphics format.")
 
  Exit Sub
'-----------------------------------------------------
PrintCustom6:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  MultiBY$ = CStr(TownRec.AppPct)
  ReportFile$ = "BLRPTS\ARAPP6.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow6
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow6
    Else
      If CustRec.VALID > XDate Then GoTo NotNow6
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm6
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm6
      End If
    End If
    GoTo NotNow6
PrintForm6:
    '                            0
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm;
    '                            1
    Print #RptHandle, QPTrim$(TownRec.AppAdd1); dlm;
    '                            2
    Print #RptHandle, QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip); dlm;
    '                            3
    Print #RptHandle, "APPLICATION FOR BUSINESS LICENSE FOR YEAR " + QPTrim$(YrUpDown(1)); dlm;
    '                            4
    Print #RptHandle, CustRec.BillName; dlm;
    '                            5
    Print #RptHandle, CustRec.ADDRESS1; dlm;
    '                            6
    Print #RptHandle, CustRec.ADDRESS2; dlm;
    '                            7
    Print #RptHandle, QPTrim$(CustRec.City) + ", " + CustRec.State + " " + CustRec.ZipCode; dlm;
    '                            8
    Print #RptHandle, "To   engage   in   business   or   profession,   make   a   separate   application"; dlm;
    '                            9
    Print #RptHandle, "for   each   business   and   each   location.  Send   fee   with   application   to"; dlm;
    '                           10
    Print #RptHandle, "The   " + QPTrim$(TownRec.AppTownOf) + " :"; dlm;
    '                      11
    Print #RptHandle, "Owners  Name: "; dlm;
    '                      12
    Print #RptHandle, "Business  Description: "; dlm;
    '                      13
    Print #RptHandle, "Business  Phone: "; dlm;
    '                      14
    Print #RptHandle, "Federal  ID  Number:"; dlm;
    '                      15
    Print #RptHandle, "State  ID  Number:"; dlm;
    '                      16
    Print #RptHandle, "To  calculate  your  " + QPTrim$(TownRec.AppTownOf) + "  Business  License  Fee,  Use  the "; dlm;
    '                      17
    Print #RptHandle, "formula  below."; dlm;
    '                      18
    Print #RptHandle, "1.  Gross  Sales"; dlm;
    '                      19
    Print #RptHandle, "2.  Less  Base  Amount"; dlm;
    '                      20
    Print #RptHandle, "3.  Excess  Gross"; dlm;
    '                      21
    Print #RptHandle, "4.  Base  Rate  Fee"; dlm;
    '                      22
    Print #RptHandle, "5.  If  No.  3  is  Greater  than"; dlm;
    '                      23
    Print #RptHandle, "Zero,  divide  No.  3  by  1,000"; dlm;
    '                      24
    Print #RptHandle, "and  round  UP"; dlm;
    '                      25
    Print #RptHandle, "6.  Multiply  #5  by  " + MultiBY$; dlm;
    '                      26
    Print #RptHandle, "7.  Total  License  Fee  # 4  +  # 6"; dlm;
    '                      27
    Print #RptHandle, "8.  Add   penalty (" + QPTrim$(Using("$##0.00", TownRec.AppColFee)) + "  Collector's"; dlm;
    If PenAmt = False Then
      '                      28
      Print #RptHandle, "Fee  and  " + QPTrim$(Using("#0.00%", TownRec.AppGrsPct / 100)) + "  per  month  after"; dlm;
    Else
      '                      28
      Print #RptHandle, "Fee  and  " + QPTrim$(Using("$##,##0.00", TownRec.AppGrsPct)) + "  per  month  after"; dlm;
    End If
    '                      29
    Print #RptHandle, QPTrim$(TownRec.AppLicRetMonth) + "  " + CStr(TownRec.AppLicRetDay); dlm;
    '                      30
    Print #RptHandle, "9.  TOTAL  DUE (# 7  +  # 8)"; dlm;
    '                      31
    Print #RptHandle, "This   is   to   certify   that   the   amount   of   total   gross   for   the   business"; dlm;
    '                      32
    Print #RptHandle, "transacted   at   or   through   the   above   location   for   the   calendar   year"; dlm;
    '                      33
    Print #RptHandle, "ending   " + QPTrim$(TownRec.AppPenMonth) + "   " + CStr(TownRec.AppPenDay) + ",   or   the   last   complete   fiscal   year   is   true   and"; dlm;
    '                      34
    Print #RptHandle, "correct,   and   that   this   report   corresponds   with   the   amount   that   was"; dlm;
    '                      35
    Print #RptHandle, "reported  to  the  SC  Tax  Commission  or  Insurance  Commission  and  with"; dlm;
    '                      36
    Print #RptHandle, "the   Internal   Revenue   Service."; dlm;
    '                      37                                  38
    Print #RptHandle, "Firm  Name/  Individual  Signature"; dlm; "By:"
    AppCnt = AppCnt + 1
NotNow6:
  Next cnt
  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp6.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #6 processed in graphics format.")
  
  Exit Sub

'-----------------------------------------------------
PrintCustom7:

  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP7.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0

  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow7
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow7
    Else
      If CustRec.VALID > XDate Then GoTo NotNow7
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm7
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm7
      End If
    End If
    GoTo NotNow7
PrintForm7:
    '                         0
    Print #RptHandle, QPTrim$(TownRec.TownName); dlm;
    '                         1
    Print #RptHandle, "BUSINESS LICENSE APPLICATION"; dlm;
    '                         2
    Print #RptHandle, "For Year: "; QPTrim$(YrUpDown(1)); dlm;
    '                         3
    Print #RptHandle, "Please  print  or  type:"; dlm;
    '                         4                                      5
    Print #RptHandle, "Applicant  Name: " + CustRec.BillName; dlm; "Phone:"; dlm;
    '                       6                    7
    Print #RptHandle, "Trade Name: "; dlm; "FEIN  or  SS#"; dlm;
    '                       8
    Print #RptHandle, "Mailing  Address:                                                                      Physical Address:"; dlm;
    '                       9
    Print #RptHandle, "Phone:                                                                               Phone:"; dlm;
    '                      10
    Print #RptHandle, "Nature  Of  Business:"; dlm;
    '                      11
    Print #RptHandle, "Gross  receipts                                             Estimated                                           Actual"; dlm;
    '                        12
    Print #RptHandle, "for  year  ending"; dlm;
    '                        13
    Print #RptHandle, QPTrim$(TownRec.AppFiscMonth) + " " + CStr(TownRec.AppFiscDay) + ", " + QPTrim$(YrUpDown(2)); dlm;
    '                        14
    Print #RptHandle, "(Wholesalers  Only...Enter  Purchases)"; dlm;
    '                        15
    Print #RptHandle, "CONTRACTORS ONLY"; dlm;
    '                        16
    Print #RptHandle, "Please   Note:   All   contractors   must   have   valid   Workmans   Compensation   coverage"; dlm;
    '                        17
    Print #RptHandle, "in   effect   for   the   time   period   covered   by   this   license.  Failure   to   have"; dlm;
    '                        18
    Print #RptHandle, "proper   coverage   will   cause   your   license   to   be   revoked."; dlm;
    '                        19
    Print #RptHandle, "____ I   certify   that   I   am   in   compliance   with   the   provisions   of   the   Virginia"; dlm;
    '                        20
    Print #RptHandle, "Workmans   Compensation   Act,   and   I   will   notify   the   " + QPTrim$(TownRec.AppTownOf); dlm;
    '                        21
    Print #RptHandle, "if   this   coverage   lapses   during   the   period   that   this   license   is   in   effect."; dlm;
    '                        22
    Print #RptHandle, "I   hereby  swear   (or   affirm)   that   the   statements   are   true,   full   and   correct   to"; dlm;
    '                        23
    Print #RptHandle, "the   best   of   my   knowledge."; dlm;
    '                        24
    Print #RptHandle, "                                              Signature                                                                       Date      "; dlm;
    '                        25
    Print #RptHandle, "***************************************************************************************************************"; dlm;
    '                        26
    Print #RptHandle, "FOR  OFFICE  USE  ONLY"; dlm;
    '                        27
    Print #RptHandle, "Zoning  classification  approved  for  this  type  of  business"; dlm;
    '                        28
    Print #RptHandle, "Approved  by "; dlm;
    '                        29
    Print #RptHandle, "                                                         Signature                                                            Date      "
    AppCnt = AppCnt + 1
NotNow7:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp7.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #7 processed in graphics format.")
  
  Exit Sub
  
PrintCustom8:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP8.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow8
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow8
    Else
      If CustRec.VALID > XDate Then GoTo NotNow8
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm8
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT5)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm8
      End If
    End If
    GoTo NotNow8
PrintForm8:
      '                            0
      Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; '"CITY OF ATMORE"
      '                            1
      Print #RptHandle, QPTrim$(TownRec.AppMayorCouncil); dlm;
      '                            2
      Print #RptHandle, "Date: "; Date$; dlm;
      '                            3
      Print #RptHandle, "NOTICE FOR RENEWAL OF BUSINESS LICENSE FOR PERIOD ENDING: " + QPTrim$(TownRec.AppFiscMonth) + ", " + QPTrim$(YrUpDown(1)); dlm;
      '                            4
      Print #RptHandle, "Business Account # " + CStr(IdxRecs(cnt)); dlm;
      '                            5
      Print #RptHandle, CustRec.BillName; dlm;
      '                            6
      Print #RptHandle, CustRec.ADDRESS1; dlm;
      '                            7
      Print #RptHandle, CustRec.ADDRESS2; dlm;
      '                            8
      Print #RptHandle, RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
      '                            9
      Print #RptHandle, "Code         Type of License"; dlm;
      
'-----------------------------------------------------------
      CatCode$ = QPTrim$(CustRec.BILLCAT1)
      GoSub GetCode
      If Len(QPTrim$(CustRec.BILLCAT1)) = 0 Then
        For x = 10 To 46
          Print #RptHandle, ""; dlm;
        Next x
'        GoTo EndAtmore1
        GoTo Next2
      End If
      '                       10                    11               12                13
      Print #RptHandle, CustRec.BILLCAT1; dlm; CustRec.DESC1; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
      
      If CodeType$ = "S" Then
        '                    14                   15                 16                 17
        Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
        
        If BaseAmt1# > 0 Then
          '                    18             19                20                 21
          Print #RptHandle, BaseAmt1; dlm; Revenue1; dlm; Percent1# / 100; dlm; Maximum1; dlm;
        Else
          '                 18       19       20       21
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt2# > 0 Then
          '                    22             23              24                  25
          Print #RptHandle, BaseAmt2; dlm; Revenue2; dlm; Percent2# / 100; dlm; Maximum2; dlm;
        Else
          '                 22       23       24       25
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt3# > 0 Then
          '                    26             27                28                29
          Print #RptHandle, BaseAmt3; dlm; Revenue3; dlm; Percent3# / 100; dlm; Maximum3; dlm;
        Else
          '                 26       27       28       29
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt4# > 0 Then
          '                    30             31                32                33
          Print #RptHandle, BaseAmt4; dlm; Revenue4; dlm; Percent4# / 100; dlm; Maximum4; dlm;
        Else
          '                 30       31       32       33
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt5# > 0 Then
          '                    34             35                36                 37
          Print #RptHandle, BaseAmt5; dlm; Revenue5; dlm; Percent5# / 100; dlm; Maximum5; dlm;
        Else
          '                 34       35       36       37
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt6# > 0 Then
          '                    38            39                 40                41
          Print #RptHandle, BaseAmt6; dlm; Revenue6; dlm; Percent6# / 100; dlm; Maximum6; dlm;
        Else
          '                 38       39       40       41
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
      Else
        For x = 14 To 41
          Print #RptHandle, ""; dlm;
        Next x
      End If
      
      If CodeType$ = "F" Then
        '                     42             43
        Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
      Else
        '                 42       43
        Print #RptHandle, ""; dlm; ""; dlm;
      End If
      
      If CodeType$ = "M" Then
        '                       44                 45                 46
        Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
      Else
        '                 44       45       46
        Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      End If
Next2:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then
'        For x = 47 To 194
        For x = 47 To 83
          Print #RptHandle, ""; dlm;
        Next x
'        GoTo EndAtmore1
        GoTo Next3
      End If
    
      CatCode$ = QPTrim$(CustRec.BILLCAT2)
      GoSub GetCode
      '                      47                      48                 49                 50
      Print #RptHandle, CustRec.BILLCAT2; dlm; CustRec.DESC2; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
      If CodeType$ = "S" Then
        '                     51                  52                 53               54
        Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
        If BaseAmt1# > 0 Then
          '                    55              56                57                  58
          Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
        Else
          '                 55       56       57       58
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt2# > 0 Then
          '                     59              60                  61               62
          Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
        Else
          '                 59       60       61       62
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt3# > 0 Then
          '                    63              64                 65                 66
          Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
        Else
          '                 63       64       65       66
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt4# > 0 Then
          '                    67               68                69                 70
          Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
        Else
          '                 67       68       69       70
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt5# > 0 Then
          '                    71               72                  73                74
          Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
        Else
          '                 71       72       73       74
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt6# > 0 Then
          '                     75             76                   77                 78
          Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
        Else
          '                 75       76       77       78
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
      Else
        For x = 51 To 78
          Print #RptHandle, ""; dlm;
        Next x
      End If
      
      If CodeType$ = "F" Then
        '                    79               80
        Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
      Else
        '                 79       80
        Print #RptHandle, ""; dlm; ""; dlm;
      End If
      
      If CodeType$ = "M" Then
        '                       81                 82                83
        Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
      Else
        '                 81       82       83
        Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      End If
Next3:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then
'        For x = 84 To 194
        For x = 84 To 120
          Print #RptHandle, ""; dlm;
        Next x
'        GoTo EndAtmore1
        GoTo Next4
      End If
      
      CatCode$ = QPTrim$(CustRec.BILLCAT3)
      GoSub GetCode
      '                        84                   85                 86                87
      Print #RptHandle, CustRec.BILLCAT3; dlm; CustRec.DESC3; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
      If CodeType$ = "S" Then
        '                     88                  89                 90                91
        Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
        If BaseAmt1# > 0 Then
          '                    92              93                 94                 95
          Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
        Else
          '                 92       93       94       95
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt2# > 0 Then
          '                     96             97                 98                 99
          Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
        Else
          '                 96       97       98       99
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt3# > 0 Then
          '                    100             101                 102               103
          Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
        Else
          '                 100      101      102      103
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt4# > 0 Then
          '                    104             105               106                 107
          Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
        Else
          '                    104             105               106                 107
          Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
        End If
        
        If BaseAmt5# > 0 Then
          '                    108             109               110                 111
          Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
        Else
          '                 108      109      110      111
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt6# > 0 Then
          '                    112             113               114                 115
          Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
        Else
          '                 112      113      114      115
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
      Else
        For x = 88 To 115
          Print #RptHandle, ""; dlm;
        Next x
      End If
      
      If CodeType$ = "F" Then
        '                     116            117
        Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
      Else
        '                 116      117
        Print #RptHandle, ""; dlm; ""; dlm;
      End If
      
      If CodeType$ = "M" Then
        '                      118                119                 120
        Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
      Else
        '                 118      119      120
        Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      End If

Next4:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then
'        For x = 121 To 194
        For x = 121 To 157
          Print #RptHandle, ""; dlm;
        Next x
'        GoTo EndAtmore1
        GoTo Next5
      End If
      
      CatCode$ = QPTrim$(CustRec.BILLCAT4)
      GoSub GetCode
      '                       121                  122                 123                124
      Print #RptHandle, CustRec.BILLCAT4; dlm; CustRec.DESC4; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
      If CodeType$ = "S" Then
        '                    125                 126                127               128
        Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
        If BaseAmt1# > 0 Then
          '                    129             130               131                 132
          Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
        Else
          '                 129      130      131      132
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt2# > 0 Then
          '                    133             134               135                 136
          Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
        Else
          '                 133      134      135      136
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt3# > 0 Then
          '                    137             138                139                140
          Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
        Else
          '                 137      138      139      140
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt4# > 0 Then
          '                    141             142                143                144
          Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
        Else
          '                 141      142      143      144
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt5# > 0 Then
          '                    145             146                147                148
          Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
        Else
          '                 145      146      147      148
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt6# > 0 Then
          '                    149             150                151               152
          Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
        Else
          '                 149      150      151      152
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
      Else
        For x = 125 To 152
          Print #RptHandle, ""; dlm;
        Next x
      End If
      
      If CodeType$ = "F" Then
        '                     153            154
        Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
      Else
        '                 153      154
        Print #RptHandle, ""; dlm; ""; dlm;
      End If
      
      If CodeType$ = "M" Then
        '                       155               156                 157
        Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
      Else
        '                 155      156      157
        Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      End If
Next5:
'-----------------------------------------------------------
      If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then
        For x = 158 To 194
          Print #RptHandle, ""; dlm;
        Next x
        GoTo EndAtmore1
      End If
      
      CatCode$ = QPTrim$(CustRec.BILLCAT5)
      GoSub GetCode
      '                       158                   159               160                 161
      Print #RptHandle, CustRec.BILLCAT5; dlm; CustRec.DESC5; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
      
      If CodeType$ = "S" Then
        '                    162                 163                 164             165
        Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
        If BaseAmt1# > 0 Then
          '                    166             167                168                169
          Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
        Else
          '                 166      167      168      169
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt2# > 0 Then
          '                    170             171               172                 173
          Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
        Else
          '                 170      171      172      173
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt3# > 0 Then
          '                    174             175               176                 177
          Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
        Else
          '                 174      175      176      177
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt4# > 0 Then
          '                    178             179               180                 181
          Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
        Else
          '                 178      179      180      181
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt5# > 0 Then
          '                     182            183               184                 185
          Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
        Else
          '                 182      183      184      185
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
        
        If BaseAmt6# > 0 Then
          '                    186             187                188                189
          Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
        Else
          '                 186      187      188      189
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
        End If
      Else
        For x = 162 To 189
          Print #RptHandle, ""; dlm;
        Next x
      End If
      
      If CodeType$ = "F" Then
        '                    190             191
        Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
      Else
        '                 190      191
        Print #RptHandle, ""; dlm; ""; dlm;
      End If
      
      If CodeType$ = "M" Then
        '                        192               193                  194
        Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
      Else
        '                 192      193      194
        Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
      End If

EndAtmore1:
      '                        195                            196
      Print #RptHandle, "Make Checks Payable To:"; dlm; "License Total: "; dlm;
      '                               197                   198
      Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; "Penalty: "; dlm;
      '                           199                      200
      Print #RptHandle, QPTrim$(TownRec.AppAdd1); dlm; "Interest: "; dlm;
      
      '                              201                   202                     203                      204
      Print #RptHandle, QPTrim$(TownRec.AppCity); dlm; "Issue Fee: "; dlm; TownRec.IssFee; dlm; QPTrim$(TownRec.SpareSpace); dlm;
      '                              205                         206                     207
      Print #RptHandle, QPTrim$(TownRec.AppState); dlm; QPTrim$(TownRec.AppZip); dlm; "Total Due: "; dlm;
      
      
      If PenAmt = False Then
        If TownRec.AppGrsPct = 8 Or TownRec.AppGrsPct = 11 Then
          '                           208
          Print #RptHandle, "License    renewals    are    due    " + QPTrim$(TownRec.AppStartMonth) + "  " + CStr(TownRec.AppStartDay) + " and    delinquent   after   " + QPTrim$(TownRec.AppLicRetMonth) + "   " + CStr(TownRec.AppLicRetDay) + " at   which   time   an  " + CStr(TownRec.AppGrsPct) + " %   penalty "; dlm;
        Else
          '                           208
          Print #RptHandle, "License   renewals   are   due   " + QPTrim$(TownRec.AppStartMonth) + "  " + CStr(TownRec.AppStartDay) + " and   delinquent   after   " + QPTrim$(TownRec.AppLicRetMonth) + "   " + CStr(TownRec.AppLicRetDay) + " at   which   time   a  " + CStr(TownRec.AppGrsPct) + " %   penalty "; dlm;
        End If
        
        If TownRec.AppDiscPct = 8 Or TownRec.AppDiscPct = 11 Then
          '                           209
          Print #RptHandle, "will  be  charged.  Renewals  after  " + QPTrim(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + " will  be  charged  an  " + CStr(TownRec.AppDiscPct) + " %  penalty.  If  you  have  any  questions  regarding  this"; dlm;
        Else
          '                           209
          Print #RptHandle, "will  be  charged.  Renewals  after  " + QPTrim(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + " will  be  charged  a  " + CStr(TownRec.AppDiscPct) + " %  penalty.  If  you  have  any  questions  regarding  this"; dlm;
        End If
      Else
          '                           208
          Print #RptHandle, "License   renewals   are   due   " + QPTrim$(TownRec.AppStartMonth) + "  " + CStr(TownRec.AppStartDay) + " and   delinquent   after   " + QPTrim$(TownRec.AppLicRetMonth) + "   " + CStr(TownRec.AppLicRetDay) + "  at   which  time  a  " + QPTrim$(Using("$##,##0.00", TownRec.AppGrsPct)) + "  penalty "; dlm;
          '                           209
          Print #RptHandle, "will  be  charged.  Renewals  after  " + QPTrim(TownRec.AppPenMonth) + "  " + CStr(TownRec.AppPenDay) + " will  be  charged  a  " + QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + "  penalty.  If  you  have  any  questions  regarding  this"; dlm;
      End If
      
      '                             210
      Print #RptHandle, "notice,  please  call  " + QPTrim$(TownRec.AppPhone) + "."; dlm;
      '                             211
      Print #RptHandle, "RENEWALS THAT DO NOT CONTAIN SIGNATURE AND GROSS RECEIPTS (WHERE REQUIRED) WILL NOT BE PROCESSED."; dlm;
      '                             212
      Print #RptHandle, "I CERTIFY THAT THE ABOVE INFORMATION IS CORRECT:"; dlm;
      '                             213
      Print #RptHandle, "NAME ____________________________________________     TITLE _______________________________________________"; dlm;
      '                             214
      Print #RptHandle, "SUBSCRIBED  AND  SWORN  TO  BEFORE  ME  THIS ________ DAY  OF  ________, ________."; dlm;
      '                             215
      Print #RptHandle, "NOTARY PUBLIC __________________________________________________________"

    AppCnt = AppCnt + 1
NotNow8:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp8.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #8 processed in graphics format.")
  
  Exit Sub

'-----------------------------------------------------
PrintCustom9:
  If QPTrim$(TownRec.AppYrUpDown(1)) = "Curr" Then
    YrUpDown(1) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "+1" Then
    YrUpDown(1) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(1)) = "-1" Then
    YrUpDown(1) = CStr(CInt(Year$) - 1)
  End If
  
  If QPTrim$(TownRec.AppYrUpDown(2)) = "Curr" Then
    YrUpDown(2) = Year$
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "+1" Then
    YrUpDown(2) = CStr(CInt(Year$) + 1)
  ElseIf QPTrim$(TownRec.AppYrUpDown(2)) = "-1" Then
    YrUpDown(2) = CStr(CInt(Year$) - 1)
  End If
  
  ReportFile$ = "BLRPTS\ARAPP9.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  
  For cnt = 1 To NumOfIdxRecs 'NumOfCustRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow9
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow9
    Else
      If CustRec.VALID > XDate Then GoTo NotNow9
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintForm9
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintForm9
      End If
    End If
    GoTo NotNow9
PrintForm9:
    '                        0
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; ' "TOWN OF HEMINGWAY, SOUTH CAROLINA"
    '                        1
    Print #RptHandle, "BUSINESS LICENSE APPLICATION"; dlm;
    '                        2
    Print #RptHandle, "For Year: " + QPTrim$(YrUpDown(1)); dlm;
    '                        3
    Print #RptHandle, "Business Name:   " + QPTrim$(CustRec.CustName); dlm;
    '                        4
    Print #RptHandle, "Mailing Address:   " + QPTrim$(CustRec.ADDRESS1); dlm;
    '                        5
    Print #RptHandle, " " + QPTrim$(CustRec.City) + " " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
    '                        6
    Print #RptHandle, "Business Address: "; dlm;
    '                        7
    Print #RptHandle, "Telephone Number: "; dlm;
    '                        8
    Print #RptHandle, "Type of Business: "; dlm;
    '                        9
    Print #RptHandle, "Social Security Number:"; dlm;
    '                       10
    Print #RptHandle, "Federal Identification Number: "; dlm;
    '                       11
    Print #RptHandle, "Gross Income Previous Year:"; dlm;
    '                       12
    Print #RptHandle, "License as Calculated:"; dlm;
    If PenAmt = False Then
      '                       13
      Print #RptHandle, QPTrim$(Using("##0.00", TownRec.AppDiscPct)) + "% Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"; dlm;
      '                       14
      Print #RptHandle, QPTrim$(Using("##0.00", TownRec.AppPct)) + "% Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"; dlm;
    Else
      '                       13
      Print #RptHandle, QPTrim$(Using("$##,##0.00", TownRec.AppDiscPct)) + " Discount, If Paid by " + QPTrim$(TownRec.AppDiscMonth) + " " + QPTrim$(Using("#0", TownRec.AppDiscDay)) + ":"; dlm;
      '                       14
      Print #RptHandle, QPTrim$(Using("$##,##0.00", TownRec.AppPct)) + " Penalty Per Month After " + QPTrim$(TownRec.AppPenMonth) + " " + QPTrim$(Using("#0", TownRec.AppPenDay)) + ":"; dlm;
    End If
    
    '                       15
    Print #RptHandle, "TOTAL AMOUNT DUE: "; dlm;
    '                       16
    Print #RptHandle, "   This   is   to   certify   that   the   above   is   a   true   statement   of   the   business"; dlm;
    '                       17
    Print #RptHandle, "transacted   at   or   through   the   above   location   for   the   calendar   year   ending"; dlm;
    '                       18
    Print #RptHandle, QPTrim$(TownRec.AppFiscMonth) + "  " + QPTrim$(Using("#0", TownRec.AppFiscDay)) + ", " + QPTrim$(YrUpDown(2)) + ",  and  that  the  report  corresponds  with  the  records  with"; dlm;
    '                       19
    Print #RptHandle, "the  S. C.  Tax  Commission  of  Insurance  Commissioner  and  with  the  Collector  of"; dlm;
    '                       20
    Print #RptHandle, "Internal  Revenue  of  the  United  States. I  understand  that  the  Town  Ordinance"; dlm;
    '                       21
    Print #RptHandle, "provides   for   penalties   of   making   false   or   fraudulent   statements   in   this"; dlm;
    '                       22
    Print #RptHandle, "application.   All   licenses   are   subject   to   being   audited.   Failure   to   provide"; dlm;
    '                       23
    Print #RptHandle, "all   information   requested   will   result   in   an   audit   from   all   required   sources."; dlm;
    '                       24
    Print #RptHandle, "Signature                                         Title                                               Date"; dlm;
    '                       25
    Print #RptHandle, "FOR OFFICE USE ONLY                                 PLEASE REMIT TO:"; dlm;
    '                       26                                    27
    Print #RptHandle, "SIC CODE____________________ "; dlm; QPTrim$(TownRec.AppTownOf); dlm; 'TOWN OF HEMINGWAY"
    '                       28                                    29
    Print #RptHandle, "RATE CLASS__________________ "; dlm; QPTrim$(TownRec.AppAdd1); dlm;             'P.O. BOX 968"
    '                       30                                    31
    Print #RptHandle, "LICENSE NUMBER_____________ "; dlm; QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip)  'HEMINGWAY S.C. 29554"
    
    AppCnt = AppCnt + 1
NotNow9:
  Next cnt

  Close         'Close all open files now

  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp9.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Application #9 processed in graphics format.")
  
  Exit Sub
'-----------------------------------------------------
PrintCustom10:
  OpenLaserFile5 LHandle
  Get LHandle, 1, Laser5
  Close LHandle
  ReportFile$ = "BLRPTS\ARAPP10.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  '                          0                               1
  Print #RptHandle, QPTrim$(Laser5.Header); dlm; MakeRegDate(Laser5.PrdBeg); dlm;
  '                             2                            3                           4
  Print #RptHandle, MakeRegDate(Laser5.PrdEnd); dlm; QPTrim$(Laser5.BLFee); dlm; Laser5.OptFeeDesc; dlm;
  
  For x = 1 To 13
    ' 5 - 17
    Print #RptHandle, Laser5.Line1(x); dlm;
  Next x
  
  For x = 1 To 10
    ' 18 - 27
    Print #RptHandle, Laser5.BusType(x); dlm;
  Next x
  For x = 1 To 10
    ' 28 - 37
    Print #RptHandle, Laser5.TaxPer(x); dlm;
  Next x
  '                           38
  Print #RptHandle, QPTrim$(Laser5.OptFee); dlm; fpcmbUseLogo.Text
  Close
  
  arBLFreeFormatApp1.Show
  frmBLLoadReport.Show
  
  MainLog ("Application #10 processed in graphics format.")
  
  Exit Sub

'-----------------------------------------------------

PrintStandard:
  CodeHandle = FreeFile
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) / Len(CodeRec)
  ReportFile$ = "BLRPTS\ARAPP1.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AppCnt = 0
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  For cnt = 1 To NumOfIdxRecs 'IdxTrNumRecs
    Get CHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotNow10
    If RangeFlag = 1 Then
      If CustRec.VALID <> XDate Then GoTo NotNow10
    Else
      If CustRec.VALID > XDate Then GoTo NotNow10
    End If
    If (CustRec.Deleted <> "Y") And QPTrim$(CustRec.SortName) <> "DELETED" And (CustRec.Inactive <> "Y") Then
      If Code$ = "ALL" Then
        GoSub PrintSTDForm
      Else
        If (InStr(1, CustRec.BILLCAT1, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT2, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT2)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT3, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT3)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT4, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT4)) = Len(QPTrim$(Code$))) _
        Or (InStr(1, CustRec.BILLCAT5, Code$) > 0 And Len(QPTrim$(CustRec.BILLCAT1)) = Len(QPTrim$(Code$))) Then _
          GoSub PrintSTDForm
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfIdxRecs 'NumOfCustRecs
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
NotNow10:
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True

  Close         'Close all open files now
  If AppCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for an advance letter using the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  Else
    arBLApp1.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("Standard application processed in graphics format.")
  
  Exit Sub

PrintSTDForm:
  GoSub GetCustFee
  NumOfCats = 0
  AppCnt = AppCnt + 1
  '                            0
  Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
  '                            1
  Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm;
  '                            2
  Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
  '                             3
  Print #RptHandle, QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
  '                             4
  Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm;
  '                             5
  Print #RptHandle, QPTrim$(CustRec.DESC1); dlm;
  '                                6
  Print #RptHandle, CStr(FeeAmt1#); dlm;
'  SCnt = 24
  If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then
    For x = 1 To 12
      Print #RptHandle, ""; dlm;
    Next x
    NumOfCats = 1
    GoTo ExitFormPrint
  End If
  '                            7
  Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm;
  '                            8
  Print #RptHandle, QPTrim$(CustRec.DESC2); dlm;
  '                        9
  Print #RptHandle, FeeAmt2#; dlm;
  If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then
    For x = 1 To 9
      Print #RptHandle, ""; dlm;
    Next x
    NumOfCats = 2
    GoTo ExitFormPrint
  End If
  '                           10
  Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm;
  '                           11
  Print #RptHandle, QPTrim$(CustRec.DESC3); dlm;
  '                      12
  Print #RptHandle, FeeAmt3#; dlm;
  
  If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then
    For x = 1 To 6
      Print #RptHandle, ""; dlm;
    Next x
    NumOfCats = 3
    GoTo ExitFormPrint
  End If
  '                             13
  Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm;
  '                             14
  Print #RptHandle, QPTrim$(CustRec.DESC4); dlm;
  '                       15
  Print #RptHandle, FeeAmt4#; dlm;
  
  If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then
    For x = 1 To 3
      Print #RptHandle, ""; dlm;
    Next x
    NumOfCats = 4
    GoTo ExitFormPrint
  End If
  '                          16
  Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm;
  '                          17
  Print #RptHandle, QPTrim$(CustRec.DESC5); dlm;
  '                       18
  Print #RptHandle, FeeAmt5#; dlm;
  NumOfCats = 5
ExitFormPrint:
'  TotalFees = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
  TotalFees = FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + IssFee#
  '                    19
  Print #RptHandle, TotalFees; dlm;
  TotalCust = TotalCust + 1
  Print #RptHandle, IssFee#
Return

GetCode:
  For Snt& = 1 To NumOfARCatRecs
    Get CodeHandle, Snt&, CodeRec
    If QPTrim$(CodeRec.CatCode) = CatCode$ Then
      CODEDESC$ = QPTrim$(CodeRec.CODEDESC)
      Select Case CodeRec.CodeType
      Case "F"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case "M"
        DESC1$ = "Per Each"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case Is = "S"
        BaseAmt1# = CodeRec.BaseAmt1
        Revenue1# = CodeRec.Recpt1
        Percent1# = CodeRec.Percent1
        Maximum1# = CodeRec.Maximum1
        BaseAmt2# = CodeRec.BaseAmt2
        Revenue2# = CodeRec.Recpt2
        Percent2# = CodeRec.Percent2
        Maximum2# = CodeRec.Maximum2
        BaseAmt3# = CodeRec.BaseAmt3
        Revenue3# = CodeRec.Recpt3
        Percent3# = CodeRec.Percent3
        Maximum3# = CodeRec.Maximum3
        BaseAmt4# = CodeRec.BaseAmt4
        Revenue4# = CodeRec.Recpt4
        Percent4# = CodeRec.Percent4
        Maximum4# = CodeRec.Maximum4
        BaseAmt5# = CodeRec.BaseAmt5
        Revenue5# = CodeRec.Recpt5
        Percent5# = CodeRec.Percent5
        Maximum5# = CodeRec.Maximum5
        BaseAmt6# = CodeRec.BaseAmt6
        Revenue6# = CodeRec.Recpt6
        Percent6# = CodeRec.Percent6
        Maximum6# = CodeRec.Maximum6
        CodeType$ = CodeRec.CodeType
      Case Else
        CodeType$ = "N"
      End Select
      Exit For
    End If
  Next Snt&

GotCode:
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
  CustFee# = OldRound#(CustFee# + FeeAmt3#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt4#)
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
  CustFee# = OldRound#(CustFee# + FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
'  FeeAmt# = 0
  
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintAppsRenwls", "PrintGraphics", Erl)
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

