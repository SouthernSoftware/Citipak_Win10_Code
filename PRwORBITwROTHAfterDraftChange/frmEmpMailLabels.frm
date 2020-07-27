VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpMailLabels 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Mailing Labels"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmEmpMailLabels.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6540
      Left            =   1553
      TabIndex        =   5
      Top             =   1103
      Width           =   8505
      _Version        =   196609
      _ExtentX        =   15002
      _ExtentY        =   11536
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmEmpMailLabels.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   3480
         TabIndex        =   4
         Tag             =   $"frmEmpMailLabels.frx":08E6
         Top             =   4410
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
         ColDesigner     =   "frmEmpMailLabels.frx":0992
      End
      Begin LpLib.fpCombo fpcmbLabel 
         Height          =   405
         Left            =   2760
         TabIndex        =   0
         Tag             =   $"frmEmpMailLabels.frx":0C8D
         Top             =   1605
         Width           =   4470
         _Version        =   196608
         _ExtentX        =   7885
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
         ColDesigner     =   "frmEmpMailLabels.frx":0D4B
      End
      Begin LpLib.fpCombo fpcmbStatus 
         Height          =   405
         Left            =   3480
         TabIndex        =   3
         Tag             =   $"frmEmpMailLabels.frx":1046
         Top             =   3720
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
         ColDesigner     =   "frmEmpMailLabels.frx":10F2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   495
         Left            =   1815
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmEmpMailLabels.frx":13ED
         Top             =   5610
         Width           =   1545
         _Version        =   131072
         _ExtentX        =   2725
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmEmpMailLabels.frx":14CC
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   495
         Left            =   3450
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Press 'Exit' to return to the main Customer Maintenance menu."
         Top             =   5610
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmEmpMailLabels.frx":16A8
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   495
         Left            =   5250
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmEmpMailLabels.frx":1886
         Top             =   5610
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmEmpMailLabels.frx":1921
      End
      Begin EditLib.fpText fptxtFirst 
         Height          =   390
         Left            =   4890
         TabIndex        =   1
         Top             =   2280
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
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
      Begin EditLib.fpText fptxtLast 
         Height          =   390
         Left            =   4890
         TabIndex        =   2
         Top             =   3030
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
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
      Begin VB.Label Label4 
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
         Height          =   345
         Left            =   1440
         TabIndex        =   15
         Top             =   3840
         Width           =   1890
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   14
         Top             =   2430
         Width           =   2190
      End
      Begin VB.Label Label6 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   3120
         Width           =   2190
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1830
         Top             =   315
         Width           =   4905
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Labels"
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
         Height          =   390
         Left            =   2250
         TabIndex        =   11
         Top             =   450
         Width           =   3945
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
         Left            =   2040
         TabIndex        =   10
         Top             =   4515
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label Type:"
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
         Left            =   1275
         TabIndex        =   9
         Top             =   1710
         Width           =   1350
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4065
         Left            =   480
         Top             =   1155
         Width           =   7575
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   1617
      TabIndex        =   12
      Top             =   7920
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
      Height          =   6795
      Left            =   1433
      Top             =   968
      Width           =   8775
   End
End
Attribute VB_Name = "frmEmpMailLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class


Private Sub cmdAlign_Click()
  Dim UBRpt As Integer
  Dim LType As Integer
  Dim cnt As Integer
  Dim Align$
  Dim ReportFile$
  
  On Error GoTo ERRORSTUFF
  
  ReDim OSet(1 To 4) As Integer
  
  Align$ = String$(23, "X")
  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  
  If fpcmbLabel.Text = "2) 1 X 3 1/2 1 Wide Text" Then
    MsgBox "Pitch 12 is recommended for this size mailing label."
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 3 1/2 3 Wide Text" Then
    MsgBox "Pitch 12 is recommended for this size mailing label."
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 3 1/2 4 Wide Text" Then
    MsgBox "Pitch 17 is recommended for this size mailing label."
    LType = 3
  Else
    fpcmbLabel.BackColor = &HFFFF&
    MsgBox "Please make a Label Type selection."
    fpcmbLabel.BackColor = &H80000005
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  UBRpt = FreeFile
  Open "PRLABEL.RPT" For Output As UBRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #UBRpt, Align$
    Next
    Print #UBRpt,
  Case 2
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$
    Next
    Print #UBRpt,
  Case 3
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #UBRpt,
  End Select

  Close UBRpt

  ViewPrint "PRLABEL.RPT", "Mailing Labels Alignment", True
  
  MainLog ("Mailing labels 'Align' feature used.")
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLMailLbls", "cmdAlign_Click", Erl)
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

Private Sub cmdExit_Click()
  frmEmployeeMaintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbLabel.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbLabel.Text, "Text") Then
    If Mid(fpcmbLabel.Text, 1, 1) = "2" Or Mid(fpcmbLabel.Text, 1, 1) = "3" Then
      MsgBox "Pitch 12 is recommended for this size mailing label."
    Else
      MsgBox "Pitch 17 is recommended for this size mailing label."
    End If
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim LType As Integer
  Dim EmpHandle As Integer
  Dim RptHandle As Integer
  Dim IdxHandle As Integer
  Dim IdxRec As Integer
  Dim NameIdxRec As NameSortIdxType
  Dim NumIdxRec As NumbSortIdxType
  Dim NumOfCustIdx As Integer
  Dim x As Integer
  Dim EmpRec As EmpData2Type
  Dim EmpCnt As Integer
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim EmpPCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim ReportFile$
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim NameFlag As Boolean
  Dim ValidCnt As Integer
  Dim FLen As Integer
  Dim LLen As Integer
  Dim FF$
  
  On Error GoTo ERRORSTUFF

  FF$ = Chr$(12)
  NameFlag = False
  If fpcmbLabel.Text = "2) 1 X 3 1/2 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 3 1/2 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 3 1/2 4 Wide Text" Then
    LType = 3
  Else
    fpcmbLabel.BackColor = &HFFFF&
    MsgBox "Please make a Label Type selection."
    fpcmbLabel.BackColor = &HFFFFFF
    fpcmbLabel.SetFocus
    Exit Sub
  End If

  If InStr(fpcmbPrintOrder.Text, "Name") Then
    NameFlag = True
    OpenEmpIdxLNameFile IdxHandle
    EmpCnt = LOF(IdxHandle) \ 2 'Len(NameIdxRec)
    ReDim EmpIdx(1 To EmpCnt) As Integer
    For x = 1 To EmpCnt
      Get IdxHandle, x, IdxRec
      EmpIdx(x) = IdxRec ' IdxNameRec.DataRecNum 'load array with record pointers
    Next x
    Close IdxHandle
  Else
    OpenEmpIdxNNameFile IdxHandle
    EmpCnt = LOF(IdxHandle) \ 2 'Len(NumIdxRec)
    ReDim EmpIdx(1 To EmpCnt) As Integer
    For x = 1 To EmpCnt
      Get IdxHandle, x, IdxRec
      EmpIdx(x) = IdxRec 'IdxNumbRec.DataRecNum 'load array with record pointers
    Next x
    Close IdxHandle
  End If

  OpenEmpData2File EmpHandle

  ReportFile$ = "PRMLLBLS.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ReDim ToPrint(1 To 5, 1 To 5) As String

  For cnt = 1 To EmpCnt
    Get EmpHandle, EmpIdx(cnt), EmpRec
    If Val(EmpRec.EmpNo) < Val(fptxtFirst.Text) Or Val(EmpRec.EmpNo) > Val(fptxtLast.Text) Then GoTo NextLabel
    If EmpRec.Deleted = 1 Then GoTo NextLabel
    If EmpRec.EMPTDATE > 0 Then GoTo NextLabel
    If fpcmbStatus.Text = "ALL" Then GoTo GoodCust
    If fpcmbStatus.Text <> QPTrim$(EmpRec.EMPSTATS) Then GoTo NextLabel
GoodCust:
    EmpPCnt = EmpPCnt + 1

    If Mid(EmpRec.EmpZip, 7, 1) <> " " Then
      Zip$ = EmpRec.EmpZip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(EmpRec.EmpZip, 5)
      Zip$ = QPTrim$(Zip$)
    End If

    Select Case LType
    Case 1
'      Print #RptHandle, "Emp #" + QPTrim$(EmpRec.EmpNo)
      Print #RptHandle,
      FLen = Len(QPTrim$(EmpRec.EmpFName))
      LLen = Len(QPTrim$(EmpRec.EmpLName))
      If FLen + LLen > 23 Then
        Print #RptHandle, Mid(EmpRec.EmpFName, 1, 1) + " " + QPTrim$(EmpRec.EmpLName)
        OverCnt = OverCnt + 1
        ReDim Preserve bigName(1 To OverCnt) As String
        bigName(OverCnt) = "From:  " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + "  To:  " + Mid(EmpRec.EmpFName, 1, 1) + " " + QPTrim$(EmpRec.EmpLName)
      Else
        Print #RptHandle, QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName)
      End If
'      Print #RptHandle, Left$(QPTrim$(CustRec.BillName), 23)
      Print #RptHandle, Left$(QPTrim$(EmpRec.EmpAddr1), 23)
      If Len(QPTrim$(EmpRec.EMPADDR2)) > 0 Then
        Print #RptHandle, Left$(QPTrim$(EmpRec.EMPADDR2), 23)
        Print #RptHandle, Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
      Else
        Print #RptHandle, Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
        Print #RptHandle,
      End If
      Print #RptHandle,
      ValidCnt = ValidCnt + 1
    Case 2
      LabelCnt = LabelCnt + 1 'this requires a line to be printed
      'in columns of 3 with each column containing data gathered
      'from different customers...it also must limit the size of some
      'variables to accommodate the limitations of a mailing label's
      'size
      FLen = Len(QPTrim$(EmpRec.EmpFName))
      LLen = Len(QPTrim$(EmpRec.EmpLName))
      If FLen + LLen > 23 Then
        ToPrint(1, LabelCnt) = Mid(EmpRec.EmpFName, 1, 1) + " " + QPTrim$(EmpRec.EmpLName)
        OverCnt = OverCnt + 1
        ReDim Preserve bigName(1 To OverCnt) As String
        bigName(OverCnt) = "From:  " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + "  To:  " + ToPrint(1, LabelCnt)
      Else
        ToPrint(1, LabelCnt) = QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName)
      End If
'      ToPrint(1, LabelCnt) = "Cust #" + QPTrim$(CustRec.CustNumb) 'Str$(AcctNumber&)
'      ToPrint(2, LabelCnt) = Left(QPTrim$(CustRec.BillName), 23)
      ToPrint(2, LabelCnt) = Left(QPTrim$(EmpRec.EmpAddr1), 23)

      If Len(QPTrim$(EmpRec.EMPADDR2)) > 0 Then
          ToPrint(3, LabelCnt) = Left(QPTrim$(EmpRec.EMPADDR2), 23)
          ToPrint(4, LabelCnt) = Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
      Else
          ToPrint(4, LabelCnt) = Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
      End If

      If LabelCnt = 3 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1

    Case 3
      LabelCnt = LabelCnt + 1
'      ToPrint(1, LabelCnt) = "Cust #" + QPTrim$(CustRec.CustNumb) 'Str$(AcctNumber&)
      FLen = Len(QPTrim$(EmpRec.EmpFName))
      LLen = Len(QPTrim$(EmpRec.EmpLName))
      If FLen + LLen > 23 Then
        ToPrint(1, LabelCnt) = Mid(EmpRec.EmpFName, 1, 1) + " " + QPTrim$(EmpRec.EmpLName)
        OverCnt = OverCnt + 1
        ReDim Preserve bigName(1 To OverCnt) As String
        bigName(OverCnt) = "From:  " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + "  To:  " + ToPrint(1, LabelCnt)
      Else
        ToPrint(1, LabelCnt) = QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName)
      End If
      
'      ToPrint(2, LabelCnt) = Left(QPTrim$(CustRec.BillName), 23)
      ToPrint(2, LabelCnt) = Left(QPTrim$(EmpRec.EmpAddr1), 23)

      If Len(QPTrim$(EmpRec.EMPADDR2)) > 0 Then
        ToPrint(3, LabelCnt) = Left(QPTrim$(EmpRec.EMPADDR2), 23)
        ToPrint(4, LabelCnt) = Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
      Else
        ToPrint(4, LabelCnt) = Left$(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim(Zip$)
      End If

      If LabelCnt = 4 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3); Tab(110); ToPrint(PCnt, 4)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1
    End Select

NextLabel:
  Next

  If LType = 2 Or LType = 3 Then 'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
    If LabelCnt > 0 Then
      For PCnt = 1 To 5
        Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
      Next
      Print #RptHandle,
    End If
  End If
  PCnt = 0
  Print #RptHandle, FF$
  Close

  If ValidCnt = 0 Then
    MsgBox "There are no employees within the parameters entered on this screen. Mailing labels not printed."
'    frmBLMessageBoxJr.Label1.Caption = "There are no customers within the parameters entered on this screen. Mailing labels not printed."
'    frmBLMessageBoxJr.Label1.Top = 700
'    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If OverCnt > 0 Then
    frmMailLblsMess.Show vbModal
    DoEvents
  End If
  
  If NameFlag = True Then
    ViewPrint ReportFile$, "Mailing Labels Sorted by Employee Name", True
  Else
    ViewPrint ReportFile$, "Mailing Labels Sorted by Employee Number", True
  End If

  KillFile ReportFile$

  MainLog ("Employee mailing labels processed in text format.")

  Exit Sub

ERRORSTUFF:
'   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpMailLabels", "PrintText", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
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
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
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
      MainLog ("Payroll.exe terminated via menu bar on frmEmpMailLabels.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim x As Integer
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim EmpCnt As Integer
  
  On Error Resume Next
  
  fpcmbStatus.Text = "ALL"
  fpcmbStatus.AddItem "ALL"
  fpcmbStatus.AddItem "Full-Time"
  fpcmbStatus.AddItem "Part-Time"
  fpcmbStatus.AddItem "Seasonal"
  fpcmbStatus.AddItem "Temporary"
  
  OpenEmpData2File EHandle
  EmpCnt = LOF(EHandle) / Len(EmpRec)
  For x = 1 To EmpCnt
    Get EHandle, x, EmpRec
    If EmpRec.Deleted = "N" Then
      fptxtFirst.Text = QPTrim$(EmpRec.EmpNo)
      Exit For
    End If
  Next x
  
  For x = EmpCnt To 1 Step -1
    Get EHandle, x, EmpRec
    If EmpRec.Deleted = "N" Then
      fptxtLast.Text = QPTrim$(EmpRec.EmpNo)
      Exit For
    End If
  Next x
  
  Close
  
  fpcmbPrintOrder.Text = "Employee Name Order"
  fpcmbPrintOrder.AddItem "Employee Name Order"
  fpcmbPrintOrder.AddItem "Employee Number Order"
  
  cmdAlign.Enabled = False
  fpcmbLabel.Text = "1) 1 X 3 1/2 3 Wide Graphical"
  fpcmbLabel.AddItem "1) 1 X 3 1/2 3 Wide Graphical"
  fpcmbLabel.AddItem "2) 1 X 3 1/2 1 Wide Text"
  fpcmbLabel.AddItem "3) 1 X 3 1/2 3 Wide Text"
  fpcmbLabel.AddItem "4) 1 X 3 1/2 4 Wide Text"
End Sub

Private Sub fpcmbLabel_Change()
  If QPTrim$(fpcmbLabel.Text) = "" Then
    fpcmbLabel.Text = "1) 1 X 3 1/2 3 Wide Graphical"
  End If
  
  If InStr(fpcmbLabel.Text, "Text") Then
    cmdAlign.Enabled = True
  Else
    cmdAlign.Enabled = False
  End If
End Sub

Private Sub fpcmbLabel_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLabel.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLabel.ListIndex = -1
  End If
  If fpcmbLabel.ListDown <> True Then
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

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLabel.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics()
  Dim IdxNameRec As NameSortIdxType
  Dim IdxNumbRec As NumbSortIdxType
  Dim IdxHandle As Integer
  Dim IdxRec As Integer
  Dim RptHandle As Integer
  Dim x As Integer
  Dim EmpRec As EmpData2Type
  Dim EmpHandle As Integer
  Dim EmpCnt As Integer
  Dim Zip$
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim ReportFile$
  Dim EmpName As String * 23
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim dlm$
  Dim ValidCnt As Integer
  Dim FLen As Integer
  Dim LLen As Integer
  
  On Error GoTo ERRORSTUFF
  
  If InStr(fpcmbPrintOrder.Text, "Name") Then
    OpenEmpIdxLNameFile IdxHandle
    EmpCnt = LOF(IdxHandle) \ Len(IdxRec)
    ReDim EmpIdx(1 To EmpCnt) As Integer
    For x = 1 To EmpCnt
      Get IdxHandle, x, IdxRec
      EmpIdx(x) = IdxRec ' IdxNameRec.DataRecNum 'load array with record pointers
    Next x
    Close IdxHandle
  Else
    OpenEmpIdxNNameFile IdxHandle
    EmpCnt = LOF(IdxHandle) \ Len(IdxRec)
    ReDim EmpIdx(1 To EmpCnt) As Integer
    For x = 1 To EmpCnt
      Get IdxHandle, x, IdxRec
      EmpIdx(x) = IdxRec 'IdxNumbRec.DataRecNum 'load array with record pointers
    Next x
    Close IdxHandle
  End If
  
  dlm = "~"
  
  OpenEmpData2File EmpHandle
  
  ReportFile$ = "PRRPTS\PRLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  ReDim ToPrint(1 To 5, 1 To 5) As String
  OverCnt = 0
  ReDim bigName(1 To 1) As String
  
  For cnt = 1 To EmpCnt
    Get EmpHandle, EmpIdx(cnt), EmpRec
    If Val(EmpRec.EmpNo) < Val(fptxtFirst.Text) Or Val(EmpRec.EmpNo) > Val(fptxtLast.Text) Then GoTo NextLabel
    If EmpRec.Deleted = 1 Then GoTo NextLabel
    If EmpRec.EMPTDATE > 0 Then GoTo NextLabel
    If fpcmbStatus.Text = "ALL" Then GoTo GoodCust
    If fpcmbStatus.Text <> QPTrim$(EmpRec.EMPSTATS) Then GoTo NextLabel
GoodCust:
    ValidCnt = ValidCnt + 1

    If Mid(EmpRec.EmpZip, 7, 1) <> " " Then
      Zip$ = EmpRec.EmpZip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(EmpRec.EmpZip, 5)
      Zip$ = QPTrim$(Zip$)
    End If
    LabelCnt = LabelCnt + 1 'this requires a line to be printed
    'in columns of 3 with each column containing data gathered
    'from different customers...it also must limit the size of some
    'variables to accommodate the limitations of a mailing label's
    'size
    FLen = Len(QPTrim$(EmpRec.EmpFName))
    LLen = Len(QPTrim$(EmpRec.EmpLName))
    If FLen + LLen > 23 Then
      ToPrint(1, LabelCnt) = Mid(EmpRec.EmpFName, 1, 1) + " " + QPTrim$(EmpRec.EmpLName)
      OverCnt = OverCnt + 1
      ReDim Preserve bigName(1 To OverCnt) As String
      bigName(OverCnt) = "From:  " + QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName) + "  To:  " + ToPrint(1, LabelCnt)
    Else
      ToPrint(1, LabelCnt) = QPTrim$(EmpRec.EmpFName) + " " + QPTrim$(EmpRec.EmpLName)
    End If
    ToPrint(2, LabelCnt) = Left(QPTrim$(EmpRec.EmpAddr1), 23)
    ToPrint(3, LabelCnt) = Left(QPTrim$(EmpRec.EMPADDR2), 23)
    ToPrint(4, LabelCnt) = Left(QPTrim$(EmpRec.EmpCity), 18) + ", " + QPTrim$(EmpRec.EmpState) + " " + QPTrim$(Zip$)
    
    If LabelCnt = 3 Then 'got a complete line
      For PCnt = 1 To 4
        '                       0                    1                      2
        Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
        ToPrint(PCnt, 1) = ""
        ToPrint(PCnt, 2) = ""
        ToPrint(PCnt, 3) = ""
        ToPrint(PCnt, 4) = ""
'        ToPrint(PCnt, 5) = ""
      Next
      Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)
      ToPrint(PCnt, 1) = ""
      ToPrint(PCnt, 2) = ""
      ToPrint(PCnt, 3) = ""
      ToPrint(PCnt, 4) = ""
'      ToPrint(PCnt, 5) = ""
      LabelCnt = 0
    End If

NextLabel:
  Next

  'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
  For PCnt = 1 To 4
    '                   0,3,6,9,12            1,4,7,10,13            2,5,8,11,14
    Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
  Next
  Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)
  
  PCnt = 0

  Close
  
  If ValidCnt = 0 Then
    MsgBox "No active employees fall within the parameters entered."
    Close
    Exit Sub
  End If
  
  If OverCnt > 0 Then
    frmMailLblsMess.Show vbModal
    DoEvents
  End If
  
  arEmpMailLabels.Show
  
  MainLog ("Mailing labels processed in graphics format.")
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmEmpMailLabels", "PrintGraphics", Erl)
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


Private Sub fpcmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStatus.ListIndex = -1
  End If
  If fpcmbStatus.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbStatus.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub
