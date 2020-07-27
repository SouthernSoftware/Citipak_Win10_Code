VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpDataPrint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Employee Data File"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmEmpDataPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5004
      Left            =   2112
      TabIndex        =   4
      Top             =   1944
      Width           =   7404
      _Version        =   196609
      _ExtentX        =   13060
      _ExtentY        =   8826
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.77
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
      Picture         =   "frmEmpDataPrint.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3555
         TabIndex        =   3
         Top             =   3405
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         ColDesigner     =   "frmEmpDataPrint.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbEmpType 
         Height          =   405
         Left            =   3555
         TabIndex        =   8
         Top             =   2925
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         ColDesigner     =   "frmEmpDataPrint.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbEmpStatus 
         Height          =   405
         Left            =   3555
         TabIndex        =   10
         Top             =   2445
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         ColDesigner     =   "frmEmpDataPrint.frx":0ED4
      End
      Begin EditLib.fpText fptxtFirstEmpNo 
         Height          =   396
         Left            =   4224
         TabIndex        =   1
         Top             =   1344
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
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
      Begin EditLib.fpText fptxtLastEmpNo 
         Height          =   396
         Left            =   4224
         TabIndex        =   2
         Top             =   1824
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
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
         Left            =   4200
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate a detailed employee personal data report."
         Top             =   4032
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
         ButtonDesigner  =   "frmEmpDataPrint.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1320
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4032
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
         ButtonDesigner  =   "frmEmpDataPrint.frx":13E2
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Employee Status:"
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
         Left            =   1392
         TabIndex        =   11
         Top             =   2544
         Width           =   1884
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Employee Type:"
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
         Left            =   1392
         TabIndex        =   9
         Top             =   3024
         Width           =   1884
      End
      Begin VB.Label Label4 
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
         Left            =   1776
         TabIndex        =   7
         Top             =   3510
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1536
         Top             =   384
         Width           =   4428
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Employee No:"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1968
         Width           =   2268
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Employee No:"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1488
         Width           =   2124
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Data File"
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
         TabIndex        =   0
         Top             =   576
         Width           =   4044
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5268
      Left            =   1980
      Top             =   1800
      Width           =   7692
   End
End
Attribute VB_Name = "frmEmpDataPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
'this report can come from either the Employee Maintenance
'Menu or the Reports Processing Menu...the "roOn" tells
'the program which menu to return to upon exit
  If frmReportsProcessing.Selection = roOn Then 'added 8/24
    frmReportsProcessing.Show
    DoEvents
    Unload frmEmpDataPrint
  Else
    frmEmployeeMaintMenu.Show
    DoEvents
    Unload frmEmpDataPrint
  End If
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpDataPrint.")
      Call Terminate
      End
    End If
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
  MainLog ("Employee Earnings Data File screen accessed.")
  Me.HelpContextID = hlpPrintEmployeeData
  Call SetupEmpDataForm
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub PrintGraphics()
  Dim RptName As String, EDistAmt(1 To 8) As Double
  Dim IdxNumOfRecs As Integer
  Dim Emp2RecLen As Integer, ECnt As Integer
  Dim DataFileSize As Long, cnt As Integer
  Dim DataNumOfRecs As Long, EDistAcct(1 To 8) As String
  Dim RptHandle As Integer
  Dim RptTitle As String
  Dim ThisSort() As Integer
  Dim FF As String, x As Integer
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim LNameIdx() As Integer
  Dim NNameIdx() As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim EmpIdxNNameHandle As Integer
  Dim EmpIdxNNameCnt As Integer
  Dim EmpIdxLNameCnt As Integer
  Dim TDate$, HDate$, BDate$
  Dim RDate$, Nextx As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  ReDim Emp2Data(1) As EmpData2Type
  ReDim Desc(1 To 53) As String * 8
  ReDim DedCodeRec(1 To 50) As DedCodeRecType
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FirstEmp&, LastEmp&
  Dim dlm$
  Dim UHandle As Integer
  Dim UnitRec As UnitFileRecType
  Dim ThisCnt As Integer
  Dim EmpZip$, ZipLen As Integer
  
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  InFileNames(4) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtLastEmpNo.Text)
  
  If LastEmp& < FirstEmp& Then
    MsgBox "The Last Employee Number is less than the First Employee Number."
    fptxtFirstEmpNo.SetFocus
    Exit Sub
  End If
  
  MaxLines = 57
  FF$ = Chr$(12)
  RptName$ = "PRRPTS\EMPDATAG.RPT"

  OpenDedCodeFile DedCodeFileHandle
  For cnt = 1 To 50
    Get DedCodeFileHandle, cnt, DedCodeRec(cnt)
    Desc$(cnt) = QPTrim$(DedCodeRec(cnt).DCDESC1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close DedCodeFileHandle
  OpenErnCodeFile ErnCodeFileHandle
  For cnt = 51 To 53
    Get ErnCodeFileHandle, cnt - 50, ErnCodeRec
    Desc$(cnt) = QPTrim$(ErnCodeRec.ERNCODE1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close ErnCodeFileHandle
 
  RptTitle$ = "Employee Information Listing"

  OpenEmpIdxNNameFile EmpIdxNNameHandle
  EmpIdxNNameCnt = LOF(EmpIdxNNameHandle) / 2
  ReDim ThisSort(EmpIdxNNameCnt)
  ReDim NNameIdx(EmpIdxNNameCnt)
  If EmpIdxNNameCnt = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Employee Data Files Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To EmpIdxNNameCnt
     Get EmpIdxNNameHandle, x, ThisSort(x)
     NNameIdx(x) = ThisSort(x)
  Next x
  
  Close EmpIdxNNameHandle
  IdxNumOfRecs = EmpIdxNNameCnt
  
  RptHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RptHandle
  OpenEmpData2File EmpData2FileHandle
  ReDim ThisSort(IdxNumOfRecs)
  For cnt = 1 To IdxNumOfRecs
    Get EmpData2FileHandle, NNameIdx(cnt), EmpData2FileRec
    'Filter
    
    If fpcmbEmpType.Text = "ALL" Then
      GoTo AllsGood
    ElseIf fpcmbEmpType.Text = "Full-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Full-Time" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Part-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Part-Time" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Seasonal" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Seasonal" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Temporary" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Temporary" Then
        GoTo SkipIt
      End If
    End If
    
AllsGood:
    If fpcmbEmpStatus.Text = "ALL" Then
      GoTo AllsGoodAgain
    ElseIf fpcmbEmpStatus.Text = "Active" Then
      If EmpData2FileRec.EMPTDATE <> 0 Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpStatus.Text = "Terminated" Then
      If EmpData2FileRec.EMPTDATE = 0 Then
        GoTo SkipIt
      End If
    End If
    
AllsGoodAgain:
    If EmpData2FileRec.EmpNo < FirstEmp Or EmpData2FileRec.EmpNo > LastEmp Then GoTo SkipIt
    If CheckValDate(Format(DateAdd("d", (EmpData2FileRec.EMPRDATE), "12-31-1979"), "mm/dd/yyyy")) = False Then
       EmpData2FileRec.EMPRDATE = 0
    End If
    For ECnt = 1 To 8
       EDistAcct(ECnt) = EmpData2FileRec.EDist(ECnt).DAcct
       EDistAmt(ECnt) = EmpData2FileRec.EDist(ECnt).DAmt
    Next ECnt
    If Not EmpData2FileRec.Deleted Then
      GoSub PrintEmpData
    End If
SkipIt:
    FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  Close EmpData2FileHandle
  Close RptHandle
  MainLog ("Employee Data File processed.")
  EnableCloseButton Me.hwnd, True
  If ThisCnt = 0 Then
    MsgBox "There are no employees listed for this criteria."
    fptxtFirstEmpNo.SetFocus
    Exit Sub
  End If
  arEmpDataRpt.Show
  frmLoadingRpt.Show

Exit Sub

PrintEmpData:
ThisCnt = ThisCnt + 1
EmpZip = QPTrim$(EmpData2FileRec.EmpZip) '06/08/04
ZipLen = Len(EmpZip) '06/08/04
If ZipLen > 5 Then
  EmpZip = Mid(EmpZip, 1, 5) + "-" + Mid(EmpZip, 6, ZipLen) '06/08/04
  EmpData2FileRec.EmpZip = EmpZip '06/08/04
End If

BDate$ = MakeRegDate(EmpData2FileRec.EMPBDAY)
If BDate = "12/31/1979" Then BDate = "No record"
HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
If HDate = "12/31/1979" Then HDate = "No record"
RDate = MakeRegDate(EmpData2FileRec.EMPRDATE)
If RDate = "12/31/1979" Then RDate = "No record"
TDate = MakeRegDate(EmpData2FileRec.EMPTDATE)
If TDate = "12/31/1979" Then TDate = "No record"
'                            0                                         1                                         2
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpNo); dlm; QPTrim$(EmpData2FileRec.EmpLName) & ", " & QPTrim$(EmpData2FileRec.EmpFName); dlm; QPTrim$(EmpData2FileRec.EmpSSN); dlm;
'                            3                                         4                                         5
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpAddr1); dlm; ""; dlm; QPTrim$(EmpData2FileRec.EMPADDR2); dlm;
'                            6                                         7                                         8
Print #RptHandle, QPTrim$(EmpData2FileRec.EmpCity); dlm; QPTrim$(EmpData2FileRec.EmpState); dlm; QPTrim$(EmpData2FileRec.EmpZip); dlm;
'                   9                        10                                       11                                      12
Print #RptHandle, BDate$; dlm; QPTrim$(EmpData2FileRec.EMPGENDR); dlm; QPTrim$(EmpData2FileRec.EMPRACE); dlm; QPTrim$(EmpData2FileRec.EMPRETNO); dlm;
'                               13                                     14                                      15                                      16
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPRETTP); dlm; QPTrim$(EmpData2FileRec.DRAFTCOD); dlm; QPTrim$(EmpData2FileRec.EMPDDACC); dlm; QPTrim$(EmpData2FileRec.PRENOTED); dlm;
'                               17                                     18                                      19                                   20
Print #RptHandle, QPTrim$(EmpData2FileRec.BankName); dlm; QPTrim$(EmpData2FileRec.BANKLOC); dlm; QPTrim$(EmpData2FileRec.TRANSIT); dlm; QPTrim$(EmpData2FileRec.EMPJOB); dlm;
'                               21                                     22                                              23                                           24
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPWCCLS); dlm; QPTrim$(EmpData2FileRec.EMPSTATS); dlm; Using("##0.00", EmpData2FileRec.EMPBCODE); dlm; QPTrim$(EmpData2FileRec.EMPPTYPE); dlm;
'                               25                                     26                                                    27                                   28
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPPFREQ); dlm; Using("##,##0.00", EmpData2FileRec.EMPPRATE); dlm; Using("##,##0.00", EmpData2FileRec.EMPORATE); dlm; HDate; dlm;
'                   29          30                       31                                     32                                        33                                          34
Print #RptHandle, RDate; dlm; TDate; dlm; QPTrim$(EmpData2FileRec.EMPFEDX); dlm; QPTrim$(EmpData2FileRec.EMPFEDO2); dlm; Using("##0.00", EmpData2FileRec.EMPFEDO1); dlm; QPTrim$(EmpData2FileRec.EMPFEDS); dlm;
'                           35                                36                                               37                                      38
Print #RptHandle, EmpData2FileRec.EMPFEDA; dlm; Using("##0.00", EmpData2FileRec.EMPFEDAA); dlm; QPTrim$(EmpData2FileRec.EMPSTAX); dlm; QPTrim$(EmpData2FileRec.EMPSTAO2); dlm;
'                           39                                                40                                   41                                     42
Print #RptHandle, Using("##0.00", EmpData2FileRec.EMPSTAO1); dlm; QPTrim$(EmpData2FileRec.EMPSTAS); dlm; EmpData2FileRec.EMPSTAA; dlm; Using("##0.00", EmpData2FileRec.EMPSTAAA); dlm;
'                           43                                        44                                           45
Print #RptHandle, QPTrim$(EmpData2FileRec.EMPSOCX); dlm; QPTrim$(EmpData2FileRec.EMPMEDX); dlm; QPTrim$(EmpData2FileRec.EMPEIC); dlm;
'46 - 245
For x = 1 To 50
  If QPTrim$(EmpData2FileRec.EmpDed(x).DPct) = "PERCENT" Then EmpData2FileRec.EmpDed(x).DPct = "PERCNT"
  Print #RptHandle, Desc$(x); dlm; EmpData2FileRec.EmpDed(x).DPct; dlm; Using$("###0.00", EmpData2FileRec.EmpDed(x).DAmt); dlm; QPTrim$(EmpData2FileRec.EmpDed(x).DOTI); dlm;
Next x

'                   246                           247                             248
Print #RptHandle, Desc(51); dlm; QPTrim$(EmpData2FileRec.EMPEACT1); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT1); dlm;
'                   249                           250                             251
Print #RptHandle, Desc(52); dlm; QPTrim$(EmpData2FileRec.EMPEACT2); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT2); dlm;
'                   252                           253                             254
Print #RptHandle, Desc(53); dlm; QPTrim$(EmpData2FileRec.EMPEACT3); dlm; Using$("###0.00", EmpData2FileRec.EMPEAMT3); dlm;
' 255 - 270
For x = 1 To 8
  Print #RptHandle, QPTrim$(EDistAcct(x)); dlm; Using$("##0.00", EDistAmt(x)); dlm;
Next x
'                                      271                                       272                                                  273
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPVACE); dlm; Using$("##0.00", EmpData2FileRec.EMPVUSED); dlm; Using$("##0.00", EmpData2FileRec.EMPVBAL); dlm;
'                                      274                                       275                                                  276
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPSLE); dlm; Using$("##0.00", EmpData2FileRec.EMPSLUSE); dlm; Using$("##0.00", EmpData2FileRec.EMPSLBAL); dlm;
'                                      277                                       278                                                  279
Print #RptHandle, Using$("##0.00", EmpData2FileRec.EMPCTE); dlm; Using$("##0.00", EmpData2FileRec.EMPCTUSE); dlm; Using$("##0.00", EmpData2FileRec.EMPCTBAL); dlm;
'                                      280                                       281                                                  282
Print #RptHandle, Using$("##0.00", EmpData2FileRec.PERERN); dlm; Using$("##0.00", EmpData2FileRec.PerUsed); dlm; Using$("##0.00", EmpData2FileRec.PERBAL); dlm;
'                                      283                                       284                                                  285
Print #RptHandle, Using$("##0.00", EmpData2FileRec.HOLERN); dlm; Using$("##0.00", EmpData2FileRec.HolUsed); dlm; Using$("##0.00", EmpData2FileRec.HOLBAL); dlm;
'                          286                              287
Print #RptHandle, EmpData2FileRec.LeaveTbl; dlm; EmpData2FileRec.ExcludeESC; dlm; UnitRec.UFEMPR; dlm; EmpData2FileRec.YN401K; dlm; QPTrim$(EmpData2FileRec.Comment)

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Private Sub SetupEmpDataForm()
   Dim First$
   Dim Last$
'  Dim EmpData1Handle As Integer, EmpIdxLNameHandle As Integer
'  Dim EmpData1Rec As EmpData1Type
'  Dim IdxRecPointer As Integer, NumOfRecs As Integer
'
'  OpenEmpData1File EmpData1Handle
'  OpenEmpIdxNNameFile EmpIdxLNameHandle
'  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
'  If NumOfRecs = 0 Then
'    MsgBox "No records on file."
'    Close
'    Exit Sub
'  End If
'  Get #EmpIdxLNameHandle, 1, IdxRecPointer
'  Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
'
'  Get #EmpIdxLNameHandle, NumOfRecs, IdxRecPointer
'  Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
'
'  Close EmpIdxLNameHandle, EmpData1Handle
  First = "0"
  Last = "0"
  Call GetFirstAndLast(First, Last)
  fptxtFirstEmpNo.Text = First
  fptxtLastEmpNo.Text = Last
  
  fpcmbEmpStatus.Text = "ALL"
  fpcmbEmpStatus.AddItem "ALL"
  fpcmbEmpStatus.AddItem "Active"
  fpcmbEmpStatus.AddItem "Terminated"
  
  fpcmbEmpType.Text = "ALL"
  fpcmbEmpType.AddItem "ALL"
  fpcmbEmpType.AddItem "Full-Time"
  fpcmbEmpType.AddItem "Part-Time"
  fpcmbEmpType.AddItem "Seasonal"
  fpcmbEmpType.AddItem "Temporary"
  
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"

End Sub

Private Sub PrintText()
  Dim RptName As String, EDistAmt(1 To 8) As Double
  Dim IdxNumOfRecs As Integer
  Dim Emp2RecLen As Integer, ECnt As Integer
  Dim DataFileSize As Long, cnt As Integer
  Dim DataNumOfRecs As Long, EDistAcct(1 To 8) As String
  Dim RptHandle As Integer
  Dim RptTitle As String
  Dim ThisSort() As Integer
  Dim FF As String, x As Integer
  Dim EmpData2FileHandle As Integer
  Dim EmpData2FileRec As EmpData2Type
  Dim LNameIdx() As Integer
  Dim NNameIdx() As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim EmpIdxLNameRec As NameSortIdxType
  Dim EmpIdxNNameRec As NumbSortIdxType
  Dim EmpIdxNNameHandle As Integer
  Dim EmpIdxNNameCnt As Integer
  Dim EmpIdxLNameCnt As Integer
  Dim TDate$, HDate$, BDate$
  Dim RDate$, Nextx As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  ReDim Emp2Data(1) As EmpData2Type
  ReDim Desc$(1 To 53)
  ReDim DedCodeRec(1 To 50) As DedCodeRecType
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FirstEmp&, LastEmp&
  Dim ThisCnt As Integer
  Dim EmpZip$, ZipLen As Integer
  
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  InFileNames(4) = "PRDATA\PRPRNDF.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtLastEmpNo.Text)
  
  If LastEmp& < FirstEmp& Then
    MsgBox "The Last Employee Number is less than the First Employee Number."
    fptxtFirstEmpNo.SetFocus
    Exit Sub
  End If
  
  MaxLines = 57
  FF$ = Chr$(12)
  RptName$ = "PRRPTS\EMPDATA.RPT"

  OpenDedCodeFile DedCodeFileHandle
  For cnt = 1 To 50
    Get DedCodeFileHandle, cnt, DedCodeRec(cnt)
    Desc$(cnt) = QPTrim$(DedCodeRec(cnt).DCDESC1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close DedCodeFileHandle
  OpenErnCodeFile ErnCodeFileHandle
  For cnt = 51 To 53
    Get ErnCodeFileHandle, cnt - 50, ErnCodeRec
    Desc$(cnt) = QPTrim$(ErnCodeRec.ERNCODE1)
    If Len(Desc$(cnt)) = 0 Then
      Desc$(cnt) = " "
    End If
  Next
  Close ErnCodeFileHandle
 
  RptTitle$ = "Employee Information Listing"

  OpenEmpIdxNNameFile EmpIdxNNameHandle
  EmpIdxNNameCnt = LOF(EmpIdxNNameHandle) / 2
  ReDim ThisSort(EmpIdxNNameCnt)
  ReDim NNameIdx(EmpIdxNNameCnt)
  If EmpIdxNNameCnt = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Employee Data Files Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To EmpIdxNNameCnt
     Get EmpIdxNNameHandle, x, ThisSort(x)
     NNameIdx(x) = ThisSort(x)
  Next x
  
  Close EmpIdxNNameHandle
  IdxNumOfRecs = EmpIdxNNameCnt
  
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  RPTSetupPRN 1, RptHandle
  OpenEmpData2File EmpData2FileHandle
  ReDim ThisSort(IdxNumOfRecs)
  For cnt = 1 To IdxNumOfRecs
    Get EmpData2FileHandle, NNameIdx(cnt), EmpData2FileRec
    'Filter
    
    If fpcmbEmpType.Text = "ALL" Then
      GoTo AllsGood
    ElseIf fpcmbEmpType.Text = "Full-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Full-Time" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Part-Time" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Part-Time" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Seasonal" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Seasonal" Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpType.Text = "Temporary" Then
      If QPTrim$(EmpData2FileRec.EMPSTATS) <> "Temporary" Then
        GoTo SkipIt
      End If
    End If
    
AllsGood:
    If fpcmbEmpStatus.Text = "ALL" Then
      GoTo AllsGoodAgain
    ElseIf fpcmbEmpStatus.Text = "Active" Then
      If EmpData2FileRec.EMPTDATE <> 0 Then
        GoTo SkipIt
      End If
    ElseIf fpcmbEmpStatus.Text = "Terminated" Then
      If EmpData2FileRec.EMPTDATE = 0 Then
        GoTo SkipIt
      End If
    End If
    
AllsGoodAgain:
    If EmpData2FileRec.EmpNo < FirstEmp Or EmpData2FileRec.EmpNo > LastEmp Then GoTo SkipIt
    If CheckValDate(Format(DateAdd("d", (EmpData2FileRec.EMPRDATE), "12-31-1979"), "mm/dd/yyyy")) = False Then
       EmpData2FileRec.EMPRDATE = 0
    End If
    For ECnt = 1 To 8
       EDistAcct(ECnt) = EmpData2FileRec.EDist(ECnt).DAcct
       EDistAmt(ECnt) = EmpData2FileRec.EDist(ECnt).DAmt
    Next ECnt
    If Not EmpData2FileRec.Deleted Then
      GoSub PrintEmpData
    End If
SkipIt:
    FrmShowPctComp.ShowPctComp cnt, IdxNumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  RPTSetupPRN 123, RptHandle '7/24 revised 8/15
  Close EmpData2FileHandle
  Close RptHandle
  If ThisCnt = 0 Then
    MsgBox "There are no employees listed for this criteria."
    fptxtFirstEmpNo.SetFocus
    Exit Sub
  End If

  ViewPrint RptName$, RptTitle$, True
  MainLog ("Employee Data File processed.")
  EnableCloseButton Me.hwnd, True

Exit Sub

PrintEmpData:
  ThisCnt = ThisCnt + 1
  EmpZip = QPTrim$(EmpData2FileRec.EmpZip) '06/08/04
  ZipLen = Len(EmpZip) '06/08/04
  If ZipLen > 5 Then
    EmpZip = Mid(EmpZip, 1, 5) + "-" + Mid(EmpZip, 6, ZipLen) '06/08/04
    EmpData2FileRec.EmpZip = EmpZip '06/08/04
  End If

  BDate$ = MakeRegDate(EmpData2FileRec.EMPBDAY)
  If BDate = "12/31/1979" Then BDate = "Not Saved  "
  HDate = MakeRegDate(EmpData2FileRec.EMPHDATE)
  If HDate = "12/31/1979" Then HDate = "Not Saved  "
  RDate = MakeRegDate(EmpData2FileRec.EMPRDATE)
  If RDate = "12/31/1979" Then RDate = "Not Saved  "
  TDate = MakeRegDate(EmpData2FileRec.EMPTDATE)
  If TDate = "12/31/1979" Then TDate = "Not Saved  "
  Print #RptHandle, "--------------------------------------------------------------------------------------------"
  Print #RptHandle, ""
  Print #RptHandle, "                  C O N F I D E N T I A L   D A T A   F I L E"
  Print #RptHandle, ""
  Print #RptHandle, "  Employee Information"
  Print #RptHandle, "       Number: "; QPTrim$(EmpData2FileRec.EmpNo); Tab(60); "Soc Sec No: "; QPTrim$(EmpData2FileRec.EmpSSN)
  Print #RptHandle, "    Last Name: "; QPTrim$(EmpData2FileRec.EmpLName); Tab(60); "First Name: "; QPTrim$(EmpData2FileRec.EmpFName)
  Print #RptHandle, "      Address: "; QPTrim$(EmpData2FileRec.EmpAddr1)
  Print #RptHandle, "      Address: "; QPTrim$(EmpData2FileRec.EMPADDR2)
  Print #RptHandle, "         City: "; QPTrim$(EmpData2FileRec.EmpCity); Tab(40); "State: "; QPTrim$(EmpData2FileRec.EmpState); Tab(60); "Zip: "; EmpData2FileRec.EmpZip
  Print #RptHandle, "    Birthdate: "; BDate$; Tab(40); "Gender: "; QPTrim$(EmpData2FileRec.EMPGENDR); Tab(60); "Race: "; QPTrim$(EmpData2FileRec.EMPRACE)
  Print #RptHandle, "   Ret Number: "; QPTrim$(EmpData2FileRec.EMPRETNO); Tab(40); "Ret Type: "; QPTrim$(EmpData2FileRec.EMPRETTP)
  Print #RptHandle, ""
  Print #RptHandle, " Direct Deposit Information"
  Print #RptHandle, "   BankDraft Code: "; QPTrim$(EmpData2FileRec.DRAFTCOD)
  Print #RptHandle, "     Bank Acct No: "; QPTrim$(EmpData2FileRec.EMPDDACC)
  Print #RptHandle, "         Prenoted: "; QPTrim$(EmpData2FileRec.PRENOTED)
  Print #RptHandle, "        Bank Name: "; QPTrim$(EmpData2FileRec.BankName)
  Print #RptHandle, "    Bank Location: "; QPTrim$(EmpData2FileRec.BANKLOC)
  Print #RptHandle, "  Bank Transit No: "; QPTrim$(EmpData2FileRec.TRANSIT)
  Print #RptHandle, ""
  Print #RptHandle, "  Job Description"
  Print #RptHandle, "        Title: "; QPTrim$(EmpData2FileRec.EMPJOB); Tab(47); "W/C Code: "; QPTrim$(EmpData2FileRec.EMPWCCLS)
  Print #RptHandle, "       Status: "; QPTrim$(EmpData2FileRec.EMPSTATS); Tab(36); "Benefit Pct: "; Using("##0.00", EmpData2FileRec.EMPBCODE); Tab(65); "Pay Type: "; QPTrim$(EmpData2FileRec.EMPPTYPE)
  Print #RptHandle, "    Frequency: "; QPTrim$(EmpData2FileRec.EMPPFREQ); Tab(36); "Rate: "; Using("##,##0.00", EmpData2FileRec.EMPPRATE); Tab(65); "O/T  Rate: "; Using("##,##0.00", EmpData2FileRec.EMPORATE)
  Print #RptHandle, "    Hire Date: "; HDate; Tab(36); "Next Review: "; RDate; Tab(65); "Term Date: "; TDate
  Print #RptHandle, "      Comment: "; QPTrim$(EmpData2FileRec.Comment)
  Print #RptHandle, ""
  Print #RptHandle, "Tax Withholding          Fixed"
  Print #RptHandle, "             Exempt  Amt/Pct  Figure   Status   # Allowances   Addit W/H Amt"
  Print #RptHandle, "   Federal: "; Tab(16); QPTrim$(EmpData2FileRec.EMPFEDX); Tab(25); QPTrim$(EmpData2FileRec.EMPFEDO2); Tab(32); Using("##0.00", EmpData2FileRec.EMPFEDO1); Tab(42); QPTrim$(EmpData2FileRec.EMPFEDS); Tab(54); EmpData2FileRec.EMPFEDA; Tab(69); Using("##0.00", EmpData2FileRec.EMPFEDAA)
  Print #RptHandle, "   State: "; Tab(16); QPTrim$(EmpData2FileRec.EMPSTAX); Tab(25); QPTrim$(EmpData2FileRec.EMPSTAO2); Tab(32); Using("##0.00", EmpData2FileRec.EMPSTAO1); Tab(42); QPTrim$(EmpData2FileRec.EMPSTAS); Tab(54); EmpData2FileRec.EMPSTAA; Tab(69); Using("##0.00", EmpData2FileRec.EMPSTAAA)
  Print #RptHandle, "   Social Security Exempt? "; QPTrim$(EmpData2FileRec.EMPSOCX); Tab(35); "Medicare Exempt? "; QPTrim$(EmpData2FileRec.EMPMEDX); Tab(63); "EIC Code: "; QPTrim$(EmpData2FileRec.EMPEIC)
  Print #RptHandle, ""
  Print #RptHandle, "   Misc Deductions   Amt/Pct   Figure  Inc O/T  Misc Deductions   Amt/Pct   Figure  Inc O/T  "
  Print #RptHandle, "   1. "; Desc$(1); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(1).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(1).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(1).DOTI);
  Print #RptHandle, "  2. "; Desc$(2); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(2).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(2).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(2).DOTI)
  Print #RptHandle, "   3. "; Desc$(3); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(3).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(3).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(3).DOTI);
  Print #RptHandle, "  4. "; Desc$(4); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(4).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(4).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(4).DOTI)
  Print #RptHandle, "   5. "; Desc$(5); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(5).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(5).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(5).DOTI);
  Print #RptHandle, "  6. "; Desc$(6); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(6).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(6).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(6).DOTI)
  Print #RptHandle, "   7. "; Desc$(7); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(7).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(7).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(7).DOTI);
  Print #RptHandle, "  8. "; Desc$(8); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(8).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(8).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(8).DOTI)
  Print #RptHandle, "   9. "; Desc$(9); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(9).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(9).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(9).DOTI);
  Print #RptHandle, " 10. "; Desc$(10); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(10).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(10).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(10).DOTI)
  Print #RptHandle, "  11. "; Desc$(11); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(11).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(11).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(11).DOTI);
  Print #RptHandle, " 12. "; Desc$(12); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(12).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(12).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(12).DOTI)
  Print #RptHandle, "  13. "; Desc$(13); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(13).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(13).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(13).DOTI);
  Print #RptHandle, " 14. "; Desc$(14); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(14).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(14).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(14).DOTI)
  Print #RptHandle, "  15. "; Desc$(15); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(15).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(15).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(15).DOTI);
  Print #RptHandle, " 16. "; Desc$(16); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(16).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(16).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(16).DOTI)
  Print #RptHandle, "  17. "; Desc$(17); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(17).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(17).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(17).DOTI);
  Print #RptHandle, " 18. "; Desc$(18); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(18).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(18).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(18).DOTI)
  Print #RptHandle, "  19. "; Desc$(19); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(19).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(19).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(19).DOTI);
  Print #RptHandle, " 20. "; Desc$(20); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(20).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(20).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(20).DOTI)
  Print #RptHandle, "  21. "; Desc$(21); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(21).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(21).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(21).DOTI);
  Print #RptHandle, " 22. "; Desc$(22); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(22).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(22).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(22).DOTI)
  Print #RptHandle, "  23. "; Desc$(23); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(23).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(23).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(23).DOTI);
  Print #RptHandle, " 24. "; Desc$(24); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(24).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(24).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(24).DOTI)
  Print #RptHandle, "  25. "; Desc$(25); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(25).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(25).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(25).DOTI);
  Print #RptHandle, " 26. "; Desc$(26); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(26).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(26).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(26).DOTI)
  Print #RptHandle, "  27. "; Desc$(27); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(27).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(27).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(27).DOTI);
  Print #RptHandle, " 28. "; Desc$(28); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(28).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(28).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(28).DOTI)
  Print #RptHandle, "  29. "; Desc$(29); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(29).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(29).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(29).DOTI);
  Print #RptHandle, " 30. "; Desc$(30); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(30).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(30).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(30).DOTI)
  Print #RptHandle, "  31. "; Desc$(31); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(31).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(31).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(31).DOTI);
  Print #RptHandle, " 32. "; Desc$(32); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(32).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(32).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(32).DOTI)
  Print #RptHandle, "  33. "; Desc$(33); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(33).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(33).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(33).DOTI);
  Print #RptHandle, " 34. "; Desc$(34); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(34).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(34).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(34).DOTI)
  Print #RptHandle, "  35. "; Desc$(35); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(35).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(35).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(35).DOTI);
  Print #RptHandle, " 36. "; Desc$(36); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(36).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(36).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(36).DOTI)
  Print #RptHandle, "  37. "; Desc$(37); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(37).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(37).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(37).DOTI);
  Print #RptHandle, " 38. "; Desc$(38); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(38).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(38).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(38).DOTI)
  Print #RptHandle, "  39. "; Desc$(39); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(39).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(39).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(39).DOTI);
  Print #RptHandle, " 40. "; Desc$(40); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(40).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(40).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(40).DOTI)
  Print #RptHandle, "  41. "; Desc$(41); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(41).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(41).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(41).DOTI);
  Print #RptHandle, " 42. "; Desc$(42); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(42).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(42).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(42).DOTI)
  Print #RptHandle, "  43. "; Desc$(43); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(43).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(43).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(43).DOTI);
  Print #RptHandle, " 44. "; Desc$(44); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(44).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(44).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(44).DOTI)
  Print #RptHandle, "  45. "; Desc$(45); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(45).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(45).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(45).DOTI);
  Print #RptHandle, " 46. "; Desc$(46); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(46).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(46).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(46).DOTI)
  Print #RptHandle, "  47. "; Desc$(47); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(47).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(47).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(47).DOTI);
  Print #RptHandle, " 48. "; Desc$(48); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(48).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(48).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(48).DOTI)
  Print #RptHandle, "  49. "; Desc$(49); Tab(23); QPTrim$(EmpData2FileRec.EmpDed(49).DPct); Tab(31); Using$("###0.00", EmpData2FileRec.EmpDed(49).DAmt); Tab(47); QPTrim$(EmpData2FileRec.EmpDed(49).DOTI);
  Print #RptHandle, " 50. "; Desc$(50); Tab(68); QPTrim$(EmpData2FileRec.EmpDed(50).DPct); Tab(76); Using$("###0.00", EmpData2FileRec.EmpDed(50).DAmt); Tab(85); QPTrim$(EmpData2FileRec.EmpDed(50).DOTI)
  Print #RptHandle, FF$
  Print #RptHandle, "--------------------------------------------------------------------------------------------"
  Print #RptHandle, ""
  Print #RptHandle, "                  C O N F I D E N T I A L   D A T A   F I L E  (C O N T)"
  Print #RptHandle, ""
  Print #RptHandle, "  Employee Information"
  Print #RptHandle, "       Number: "; QPTrim$(EmpData2FileRec.EmpNo); Tab(60); "Soc Sec No: "; QPTrim$(EmpData2FileRec.EmpSSN)
  Print #RptHandle, "    Last Name: "; QPTrim$(EmpData2FileRec.EmpLName); Tab(60); "First Name: "; QPTrim$(EmpData2FileRec.EmpFName)
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, "   Default Earning Codes    Account Number     Earnings"
  Print #RptHandle, "   1. "; Desc(51); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT1); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT1)
  Print #RptHandle, "   2. "; Desc(52); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT2); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT2)
  Print #RptHandle, "   3. "; Desc(53); Tab(33); QPTrim$(EmpData2FileRec.EMPEACT3); Tab(48); Using$("###0.00", EmpData2FileRec.EMPEAMT3)
  Print #RptHandle, ""
  Print #RptHandle, "   Wage Account Numbers        Default Distribution"
  Print #RptHandle, "   1. "; QPTrim$(EDistAcct(1)); Tab(38); Using$("##0.00", EDistAmt(1))
  Print #RptHandle, "   2. "; QPTrim$(EDistAcct(2)); Tab(38); Using$("##0.00", EDistAmt(2))
  Print #RptHandle, "   3. "; QPTrim$(EDistAcct(3)); Tab(38); Using$("##0.00", EDistAmt(3))
  Print #RptHandle, "   4. "; QPTrim$(EDistAcct(4)); Tab(38); Using$("##0.00", EDistAmt(4))
  Print #RptHandle, "   5. "; QPTrim$(EDistAcct(5)); Tab(38); Using$("##0.00", EDistAmt(5))
  Print #RptHandle, "   6. "; QPTrim$(EDistAcct(6)); Tab(38); Using$("##0.00", EDistAmt(6))
  Print #RptHandle, "   7. "; QPTrim$(EDistAcct(7)); Tab(38); Using$("##0.00", EDistAmt(7))
  Print #RptHandle, "   8. "; QPTrim$(EDistAcct(8)); Tab(38); Using$("##0.00", EDistAmt(8))
  Print #RptHandle, ""
  Print #RptHandle, " Benefit Schedule          Earned        Used      Balance"
  Print #RptHandle, "     Vacation"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPVACE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPVUSED); Tab(54); Using$("##0.00", EmpData2FileRec.EMPVBAL)
  Print #RptHandle, "     Sick Leave"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPSLE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPSLUSE); Tab(54); Using$("##0.00", EmpData2FileRec.EMPSLBAL)
  Print #RptHandle, "     Comp Time"; Tab(29); Using$("##0.00", EmpData2FileRec.EMPCTE); Tab(42); Using$("##0.00", EmpData2FileRec.EMPCTUSE); Tab(54); Using$("##0.00", EmpData2FileRec.EMPCTBAL)
  Print #RptHandle, "     Personal"; Tab(29); Using$("##0.00", EmpData2FileRec.PERERN); Tab(42); Using$("##0.00", EmpData2FileRec.PerUsed); Tab(54); Using$("##0.00", EmpData2FileRec.PERBAL)
  Print #RptHandle, "     Holiday"; Tab(29); Using$("##0.00", EmpData2FileRec.HOLERN); Tab(42); Using$("##0.00", EmpData2FileRec.HolUsed); Tab(54); Using$("##0.00", EmpData2FileRec.HOLBAL)
  Print #RptHandle, ""
  Print #RptHandle, "     Leave Table"; Tab(20); EmpData2FileRec.LeaveTbl; Tab(30); "Exclude ESC"; Tab(45); EmpData2FileRec.ExcludeESC
  Print #RptHandle, FF$
  
  Return
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

Private Sub fpcmbEmpStatus_Click()
  If fpcmbEmpStatus.Text = "" Then
    fpcmbEmpStatus.Text = "ALL"
  End If
End Sub

Private Sub fpcmbEmpStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEmpStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEmpStatus.ListIndex = -1
  End If
  If fpcmbEmpStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbEmpType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub


Private Sub fpcmbEmpType_Click()
  If fpcmbEmpType.Text = "" Then
    fpcmbEmpType.Text = "ALL"
  End If
End Sub

Private Sub fpcmbEmpType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbEmpType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEmpType.ListIndex = -1
  End If
  If fpcmbEmpType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub GetFirstAndLast(ByRef First$, ByRef Last$)
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfRecs As Integer
  Dim XHandle As Integer
  Dim Thisx As Integer
  Dim x As Integer
  
  OpenEmpIdxNNameFile XHandle
  NumOfRecs = LOF(XHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No employee records have been saved."
    Close
    Exit Sub
  End If
  
  OpenEmpData2File EHandle
  For x = 1 To NumOfRecs
    Get XHandle, x, Thisx
    Get EHandle, Thisx, EmpRec
    If EmpRec.Deleted = -1 Then GoTo SkipD1
    If Val(EmpRec.EmpNo) > 0 Then
      First = QPTrim$(EmpRec.EmpNo)
      Exit For
    End If
SkipD1:
  Next x
  
  For x = NumOfRecs To 1 Step -1
    Get XHandle, x, Thisx
    Get EHandle, Thisx, EmpRec
    If EmpRec.Deleted = -1 Then GoTo SkipD2
    If Val(EmpRec.EmpNo) > 0 Then
      Last = QPTrim$(EmpRec.EmpNo)
      Exit For
    End If
SkipD2:
  Next x
    
End Sub
