VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLLicFormLaser 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Forms Printing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11580
   Icon            =   "frmBLLicFormLaser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7716
      Left            =   672
      TabIndex        =   13
      Top             =   576
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
      Picture         =   "frmBLLicFormLaser.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   6045
         TabIndex        =   10
         Tag             =   $"frmBLLicFormLaser.frx":08E6
         Top             =   3900
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
         ColDesigner     =   "frmBLLicFormLaser.frx":0992
      End
      Begin LpLib.fpCombo fpcmbPrintFeesYN 
         Height          =   405
         Left            =   7440
         TabIndex        =   11
         Tag             =   $"frmBLLicFormLaser.frx":0C89
         Top             =   4800
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
         ColDesigner     =   "frmBLLicFormLaser.frx":0DCA
      End
      Begin LpLib.fpCombo fpcmbBalanceType 
         Height          =   384
         Left            =   6528
         TabIndex        =   12
         Tag             =   $"frmBLLicFormLaser.frx":10C1
         Top             =   5616
         Width           =   2784
         _Version        =   196608
         _ExtentX        =   4911
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
         ColDesigner     =   "frmBLLicFormLaser.frx":12B1
      End
      Begin LpLib.fpCombo fpcmbSignature 
         Height          =   405
         Left            =   8550
         TabIndex        =   9
         Tag             =   $"frmBLLicFormLaser.frx":15A8
         Top             =   2970
         Width           =   960
         _Version        =   196608
         _ExtentX        =   1693
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
         ColDesigner     =   "frmBLLicFormLaser.frx":1696
      End
      Begin EditLib.fpText fptxtBegNum 
         Height          =   396
         Left            =   1536
         TabIndex        =   2
         Tag             =   $"frmBLLicFormLaser.frx":198D
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
         Left            =   4416
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'License Processing' menu."
         Top             =   6624
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
         ButtonDesigner  =   "frmBLLicFormLaser.frx":1C6C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   630
         Left            =   6675
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Press 'Process' to begin printing the business license forms using the parameters entered above."
         Top             =   6630
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
         ButtonDesigner  =   "frmBLLicFormLaser.frx":1E4B
      End
      Begin EditLib.fpDateTime fptxtVThru 
         Height          =   370
         Left            =   2496
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
      Begin EditLib.fpDateTime fptxtIssueDate 
         Height          =   370
         Left            =   7776
         TabIndex        =   7
         Tag             =   $"frmBLLicFormLaser.frx":202A
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
      Begin EditLib.fpText fptxtHeading 
         Height          =   396
         Index           =   0
         Left            =   864
         TabIndex        =   3
         Tag             =   $"frmBLLicFormLaser.frx":20CA
         Top             =   4416
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
         Tag             =   $"frmBLLicFormLaser.frx":21A7
         Top             =   4848
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
         Tag             =   $"frmBLLicFormLaser.frx":2285
         Top             =   5280
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
      Begin EditLib.fpText fptxtAuthorizedBy 
         Height          =   396
         Left            =   2592
         TabIndex        =   6
         Tag             =   "The person's name entered here will appear on the business license as the town offical responsible for issuing business licenses."
         Top             =   5808
         Width           =   3228
         _Version        =   196608
         _ExtentX        =   5694
         _ExtentY        =   698
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
      Begin EditLib.fpDateTime fptxtFromDate 
         Height          =   370
         Left            =   2496
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
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   636
         Left            =   1872
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   $"frmBLLicFormLaser.frx":2362
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
         ButtonDesigner  =   "frmBLLicFormLaser.frx":2432
      End
      Begin EditLib.fpDateTime fpBLYear 
         Height          =   375
         Left            =   7050
         TabIndex        =   8
         Tag             =   "The date entered here will appear on the business license as the active year for this license"
         Top             =   2445
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
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
         Height          =   345
         Left            =   3165
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   $"frmBLLicFormLaser.frx":2615
         Top             =   3405
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   609
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
         ButtonDesigner  =   "frmBLLicFormLaser.frx":2717
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Signature Line (Y/N)?"
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
         Left            =   5568
         TabIndex        =   31
         Top             =   3072
         Width           =   2796
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
         TabIndex        =   30
         Top             =   2064
         Width           =   2988
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
         Left            =   1920
         TabIndex        =   28
         Top             =   7296
         Width           =   2100
      End
      Begin VB.Label Label10 
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
         TabIndex        =   26
         Top             =   5280
         Width           =   3228
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
         Left            =   1872
         TabIndex        =   25
         Top             =   2478
         Width           =   492
      End
      Begin VB.Label Label1 
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
         Left            =   1632
         TabIndex        =   24
         Top             =   1938
         Width           =   732
      End
      Begin VB.Label Label11 
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
         Left            =   1440
         TabIndex        =   23
         Top             =   1536
         Width           =   3276
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Authorized By:"
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
         Left            =   816
         TabIndex        =   22
         Top             =   5856
         Width           =   1692
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
         TabIndex        =   21
         Top             =   3552
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Laser Business License"
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
         Left            =   2976
         TabIndex        =   20
         Top             =   432
         Width           =   4572
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   2835
         Top             =   288
         Width           =   4908
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
         Left            =   1536
         TabIndex        =   19
         Top             =   3072
         Width           =   3036
      End
      Begin VB.Label Label3 
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
         TabIndex        =   18
         Top             =   1576
         Width           =   2412
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
         TabIndex        =   17
         Top             =   4032
         Width           =   1980
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
         TabIndex        =   16
         Top             =   4464
         Width           =   2796
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   1440
      TabIndex        =   29
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
      MaxWidth        =   5000
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
      Left            =   468
      Top             =   372
      Width           =   10644
   End
End
Attribute VB_Name = "frmBLLicFormLaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim UsePermLicNum As Boolean

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtFromDate.ToolTipText = ""
    fptxtVThru.ToolTipText = ""
    fptxtIssueDate.ToolTipText = ""
    fptxtBegNum.ToolTipText = ""
    cmdList.ToolTipText = ""
    fptxtHeading(0).ToolTipText = ""
    fptxtHeading(1).ToolTipText = ""
    fptxtHeading(2).ToolTipText = ""
    fptxtAuthorizedBy.ToolTipText = ""
'    fpcmbFeeYN.ToolTipText = ""
    fpcmbPrintFeesYN.ToolTipText = ""
    fpcmbBalanceType.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    fpcmbPrintOrder.ToolTipText = ""
    cmdHelp.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtFromDate.ToolTipText = "The date entered here will appear on the license as the first valid date."
'    fptxtVThru.ToolTipText = "The date entered here will appear on the license as the expiration date."
'    fptxtIssueDate.ToolTipText = "Enter the date which will appear on the business licenses indicating the license issuance date."
'    fptxtBegNum.ToolTipText = "Enter the new business license number that will begin the license printing process."
'    cmdList.ToolTipText = "Use this button to bring up a list of all customer's and their license numbers."
'    fptxtHeading(0).ToolTipText = "Optional line of text that will appear as the first line of the license header."
'    fptxtHeading(1).ToolTipText = "Optional line of text that will appear as the second line of the license header."
'    fptxtHeading(2).ToolTipText = "Optional line of text that will appear as the third line of the license header."
'    fptxtAuthorizedBy.ToolTipText = "The name entered here will appear on the laser license as a town official."
'    fpcmbFeeYN.ToolTipText = "Select Yes and any fees calculated and not posted for these licenses will be reset to zero. Choose No to allow the unposted calculations to remain."
'    fpcmbPrintFeesYN.ToolTipText = "This option allows the current fees to appear on each license. This option is disabled if the 'Charge Account With Fee Y/N' option is No."
'    fpcmbBalanceType.ToolTipText = "Business licenses can be printed with total outstanding balances and current balances or just current balances."
'    cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'    cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  End If
End Sub

Private Sub cmdList_Click()
  frmBLLicenseNumList.Show vbModal
End Sub

Private Sub fpcmbBalanceType_Change()
  If QPTrim$(fpcmbBalanceType.Text) = "" Then
    fpcmbBalanceType.Text = "Current Balance Only"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLLicFormLaser.")
      Call Terminate
      End
    End If
  End If
End Sub

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
  Unload frmBLLicFormLaser
End Sub

Private Sub cmdProcess_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  
  OpenTownFile THandle
  Get THandle, 1, TownRec
  Close THandle
  
  If Not Exist("artmppst.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please process Business License registers first."
    frmBLMessageBoxJr.Label1.Top = 800
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
  
  If Date2Num(fptxtVThru.Text) < Date2Num(fptxtIssueDate.Text) Then
    frmBLMessageBoxJr.Label1.Caption = "The issue date comes before the new expiration date. Please revise these dates."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtVThru.SetFocus
    Exit Sub
  ElseIf Date2Num(fptxtVThru.Text) = Date2Num(fptxtIssueDate.Text) Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The issue date and the new expiration date are the same. If this is correct then press F10 to continue. Otherwise press ESC to return to the screen."
    frmBLMessageBoxJrWOpts.Label1.Top = 700
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
      Exit Sub
    End If
  End If
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  Call PrintLic
  
End Sub

Private Sub LoadMe()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim NewYear$
  Dim DHandle As Integer
  Dim ThisDate$
  Dim ThisHeader$
  
  On Error Resume Next
  lblBalloon.Visible = False
'  fptxtFromDate.ToolTipText = "The date entered here will appear on the license as the first valid date."
'  fptxtVThru.ToolTipText = "The date entered here will appear on the license as the expiration date."
'  fptxtIssueDate.ToolTipText = "Enter the date which will appear on the business licenses indicating the license issuance date."
'  fptxtBegNum.ToolTipText = "Enter the new business license number that will begin the license printing process."
'  cmdList.ToolTipText = "Use this button to bring up a list of all customer's and their license numbers."
'  fptxtHeading(0).ToolTipText = "Optional line of text that will appear as the first line of the license header."
'  fptxtHeading(1).ToolTipText = "Optional line of text that will appear as the second line of the license header."
'  fptxtHeading(2).ToolTipText = "Optional line of text that will appear as the third line of the license header."
'  fptxtAuthorizedBy.ToolTipText = "The name entered here will appear on the laser license as a town official."
'  fpcmbFeeYN.ToolTipText = "Select Yes and any fees calculated and not posted for these licenses will be reset to zero. Choose No to allow the unposted calculations to remain."
'  fpcmbPrintFeesYN.ToolTipText = "This option allows the current fees to appear on each license. This option is disabled if the 'Charge Account With Fee Y/N' option is No."
'  fpcmbBalanceType.ToolTipText = "Business licenses can be printed with total outstanding balances and current balances or just current balances."
'  cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'  cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
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
  
  If Exist("appheader.dat") Then
    DHandle = FreeFile
    Open "appheader.dat" For Input As #DHandle
    Line Input #DHandle, ThisHeader
    fptxtHeading(0).Text = ThisHeader
    Close DHandle
  Else
    fptxtHeading(0).Text = "MUNICIPAL LICENSE"
  End If
  
  UsePermLicNum = False
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
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
  fptxtIssueDate = Date
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fpcmbPrintFeesYN.Text = "Yes"
  fpcmbPrintFeesYN.AddItem "No"
  fpcmbPrintFeesYN.AddItem "Yes"
  fpcmbSignature.Text = "Yes"
  fpcmbSignature.AddItem "Yes"
  fpcmbSignature.AddItem "No"
  fpcmbBalanceType.Text = "Current Balance Only"
  fpcmbBalanceType.AddItem "Current Balance Only"
  fpcmbBalanceType.AddItem "Total Balance"
  fptxtHeading(1).Text = QPTrim$(TownRec.TownName)
  Select Case QPTrim$(TownRec.State)
    Case "NC"
      fptxtHeading(2).Text = "STATE OF NORTH CAROLINA"
    Case "SC"
      fptxtHeading(2).Text = "STATE OF SOUTH CAROLINA"
    Case "VA"
      fptxtHeading(2).Text = "STATE OF VIRGINIA"
    Case "GA"
      fptxtHeading(2).Text = "STATE OF GEORGIA"
    Case "AR"
      fptxtHeading(2).Text = "STATE OF ARKANSAS"
    Case "AL"
      fptxtHeading(2).Text = "STATE OF ALABAMA"
    Case "OK"
      fptxtHeading(2).Text = "STATE OF OKLAHOMA"
    Case Else
      fptxtHeading(2).Text = "UNKNOWN STATE"
  End Select
  fptxtAuthorizedBy.Text = QPTrim$(TownRec.Contact)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub PrintLic()
  Dim ReportFile$
  Dim x As Double, y As Integer
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
  Dim CustLicNum$
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
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim IssueDate$
  Dim SCnt As Integer, LCnt As Integer
  Dim TempHandle As Integer
  Dim TempRec As TempTransPostType
  Dim TempNum As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim PrintFees As Boolean
  Dim dlm$
  Dim One As Integer
  Dim DHandle As Integer
  Dim PCnt As Integer
  Dim ThisCat As String * 35
  Dim BalanceFlag As Integer
  Dim ThisAdd As String * 35
  Dim AddEmptyFields As Integer
  Dim Nextx As Double
  Dim ThisDate$
  Dim ThisLen As Integer
  Dim ThisHeader$
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  
  If fpcmbSignature.Text = "Yes" Then
    PrintSign = True
  Else
    PrintSign = False
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
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

  StoreExpireDate$ = fptxtVThru.Text
  ExpireDate$ = fptxtVThru.Text
  Year$ = Mid(fptxtVThru.Text, 7, 4)
  
  IssueDate$ = fptxtIssueDate.Text
  PostDate$ = fptxtIssueDate.Text
  ReportFile$ = "BLRPTS\ARLASER.RPT"
  CustCnt = 0
  
  PrintFees = False
  If QPTrim$(fpcmbPrintFeesYN.Text) = "Yes" Then
    PrintFees = True
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
    ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
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
  cmdHelp.Enabled = False
  
  frmBLShowPctComp.cmdCancel.Visible = False
  
  For x = 1 To NumOfCustIdxRecs
    For y = 1 To NumOfTempRecs
      Get TempHandle, y, TempRec
        If CDbl(TempRec.CustomerNumber) = IdxRecs(x) Then
          Nextx = Nextx + 1
          ReDim Preserve PrintIdx(1 To Nextx) As Double
          PrintIdx(Nextx) = y
          Exit For
        End If
    Next y
    frmBLShowPctComp.ShowPctComp x, NumOfCustIdxRecs
  Next x
  
  Unload frmBLShowPctComp
  
  frmBLShowPctComp.Label1 = "Printing Customer Business Licenses"
  frmBLShowPctComp.Show
  
  frmBLShowPctComp.cmdCancel.Visible = False
  
  KillFile ("licprnOK.dat")

  If InStr(fpcmbBalanceType.Text, "Only") Then
    BalanceFlag = 1
  Else
    BalanceFlag = 2
  End If
  
  If PrintFees = False Then BalanceFlag = 1
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, PrintIdx(x), TempRec
    Get CustHandle, Val(TempRec.CustomerNumber), CustRec
      If UsePermLicNum = True Then LICENSE# = Val(CustRec.LICENSE)
      CustLicNum = TempRec.LICENSE
      ThisAdd = QPTrim$(CustRec.ADDRESS1)
      '                     0                1                 2             3
      Print #RptHandle, SHeading1$; dlm; SHeading2$; dlm; SHeading3$; dlm; Year$; dlm;
      If CustRec.Prorate < 100 Then
        '                       4                        5
        Print #RptHandle, CustLicNum; dlm; CStr(CustRec.Prorate); dlm;
      Else
        '                       4               5
        Print #RptHandle, CustLicNum; dlm; ""; dlm;
      End If
      '                              6                      7
      Print #RptHandle, QPTrim$(CustRec.Contact); dlm; LICENSE#; dlm;
      '                       8                          9
      Print #RptHandle, ThisAdd; dlm; QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + "  " + QPTrim$(CustRec.ZipCode); dlm;
      '                     10               11                      12                                  13
      Print #RptHandle, IssueDate$; dlm; ExpireDate$; dlm; QPTrim$(CustRec.CustName); dlm; QPTrim$(fptxtAuthorizedBy.Text); dlm;
      
      AddEmptyFields = 0
      
      If QPTrim$(CustRec.BILLCAT1) <> "" Then
        If PrintFees = True Then
          '                             14                          15                            16
          Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCatDesc(CustRec.BILLCAT1); dlm; TempRec.CatFee1; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCatDesc(CustRec.BILLCAT1); dlm; ""; dlm;
        End If
      Else
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT2) <> "" Then
        If PrintFees = True Then
          '                             17                          18                            19
          Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCatDesc(CustRec.BILLCAT2); dlm; TempRec.CatFee2; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCatDesc(CustRec.BILLCAT2); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
        
      If QPTrim$(CustRec.BILLCAT3) <> "" Then
        If PrintFees = True Then
          '                             20                          21                            22
          Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCatDesc(CustRec.BILLCAT3); dlm; TempRec.CatFee3; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCatDesc(CustRec.BILLCAT3); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT4) <> "" Then
        If PrintFees = True Then
          '                             23                          24                            25
          Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCatDesc(CustRec.BILLCAT4); dlm; TempRec.CatFee4; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCatDesc(CustRec.BILLCAT4); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT5) <> "" Then
        If PrintFees = True Then
          '                             26                          27                            28
          Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCatDesc(CustRec.BILLCAT5); dlm; TempRec.CatFee5; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCatDesc(CustRec.BILLCAT5); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      For y = 1 To AddEmptyFields
        '
        Print #RptHandle, ""; dlm;
      Next y
      
      
      If PrintFees = True Then
        
        If OldRound(TownRec.IssFee) > 0 Then
          '                            29
          Print #RptHandle, OldRound(TownRec.IssFee); dlm;
        Else
          '                 29
          Print #RptHandle, "0"; dlm;
        End If
        
      'Calc Total License Amount Here
      Else
        '                  29
        Print #RptHandle, "0"; dlm;
      End If
        
      TotalBillAmt# = OldRound(TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TownRec.IssFee)
        
      If PrintFees = True Then
        '                      30
        Print #RptHandle, TotalBillAmt#; dlm;
      Else
        '                 30
        Print #RptHandle, "No"; dlm;
      End If
      '                         31
      Print #RptHandle, fptxtFromDate.Text; dlm;
      
      If BalanceFlag = 2 Then
        '                               32          33                    34                  35
        Print #RptHandle, CustRec.AcctBal; dlm; TempRec.AcctBal; dlm; BalanceFlag; dlm; fpBLYear.Text; dlm;
      Else
        '                 32     33          34                  35
        Print #RptHandle, 0; dlm; 0; dlm; BalanceFlag; dlm; fpBLYear.Text; dlm;
      End If
      '                             36
      Print #RptHandle, QPTrim$(CustRec.BillName); dlm; QPTrim$(CustRec.ServAdd)
      
      GoSub Post2TempAccount
      If UsePermLicNum = False Then LICENSE# = LICENSE# + 1
      frmBLShowPctComp.ShowPctComp x, NumOfTempRecs
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now

  If PCnt > 0 Then
    One = 1
    DHandle = FreeFile
    Open "licprnOK.dat" For Output As DHandle Len = 2
    Print #DHandle, One
    Close DHandle
  End If
  
  arBLLaser.Show
  frmBLLoadReport.Show
  
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
  
  MainLog ("Business license laser forms printed.")
  Exit Sub
  
Post2TempAccount:
  TempRec.LICENSE = LTrim$(Str$(LICENSE#))
  TempRec.VALID = Date2Num%(StoreExpireDate$)
  TempRec.TransDate = Date2Num%(PostDate$)
  If CustRec.FirstTrans = 0 Then
    TempRec.Prev = 0
  Else
    TempRec.Prev = CustRec.LastTrans
  End If
  
  Put TempHandle, PrintIdx(x), TempRec
  PCnt = PCnt + 1
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicFormLaser", "PrintLic", Erl)
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

Private Sub fpDateTime1_Change()

End Sub

Private Sub fpcmbSignature_Change()
  If QPTrim$(fpcmbSignature.Text) = "" Then
    fpcmbSignature.Text = "Yes"
  End If
End Sub

Private Sub fpcmbSignature_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSignature.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSignature.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
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

Private Sub fptxtAuthorizedBy_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbPrintOrder.SetFocus
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
      If (QPTrim$(CustRec.SortName) = "DELETED" Or QPTrim(CustRec.Deleted) <> "Y") And QPTrim$(CustRec.IssueLicense) = "N" Then
      If Not IsNumeric(CustRec.LICENSE) Then GoTo NoLicenseNum
        ThatLic = CDbl(CustRec.LICENSE)
        For y = 1 To YCnt
          If YCntIdx(y) = ThatLic Then
            frmBLMessageBoxJr.Label1.Caption = "The beginning license number entered would cause a duplicate license number problem between new License # " + CStr(YCntIdx(y)) + " and the existing license number of current customer " + QPTrim(CustRec.CustName) + " who is not included in this license process. Please revise your beginning license number to avoid this conflict."
            frmBLMessageBoxJr.Label1.Top = 430
            frmBLMessageBoxJr.Label1.Height = 1300
            frmBLMessageBoxJr.Show vbModal
            fptxtBegNum.SetFocus
            Look4DupLicNums = True
            GoTo NoLicenseNum
          End If
        Next y
      End If
NoLicenseNum:
  Next x
  
  Exit Function
  
ERRORSTUFF:
  frmBLMessageBoxJr.Label1.Caption = "ERROR: An error has occurred in the 'Look4DupLicNum' function for customer number " + QPTrim$(CustRec.CustNumb) + "."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  Close CustHandle
  
End Function

