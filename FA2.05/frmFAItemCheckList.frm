VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAItemCheckList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Check List Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAItemCheckList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6012
      Left            =   1944
      TabIndex        =   4
      Top             =   1428
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   10604
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAItemCheckList.frx":08CA
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         ToolTipText     =   "Select the order this report will display data."
         Top             =   1770
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
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
         ColDesigner     =   "frmFAItemCheckList.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbYN 
         Height          =   405
         Left            =   5475
         TabIndex        =   1
         ToolTipText     =   "Enter a Y if disposed of items should appear in the report. Choose N if disposed of items should be excluded from this report."
         Top             =   2400
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
         ColDesigner     =   "frmFAItemCheckList.frx":0BDD
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   3
         ToolTipText     =   "Select Graphical for a robust but slower reporting process. Choose Text for a quicker report."
         Top             =   3930
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
         Style           =   0
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
         ColDesigner     =   "frmFAItemCheckList.frx":0ED4
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   2
         ToolTipText     =   "If DEPARTMENT NUMBER is selected in Report Order then enter the department to report on here."
         Top             =   3072
         Width           =   1068
         _Version        =   196608
         _ExtentX        =   1884
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
         MaxLength       =   14
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
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4512
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   4800
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAItemCheckList.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1584
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   4800
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAItemCheckList.frx":13AA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   390
         Left            =   4410
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current departments."
         Top             =   3075
         Width           =   1365
         _Version        =   131072
         _ExtentX        =   2408
         _ExtentY        =   688
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAItemCheckList.frx":1586
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Order:"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1872
         Width           =   1548
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Include Disposed Of Items (Y/N):"
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
         TabIndex        =   8
         Top             =   2496
         Width           =   3660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dept #"
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
         Left            =   1968
         TabIndex        =   7
         Top             =   3168
         Width           =   924
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   576
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Check List Report"
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
         Height          =   375
         Left            =   1950
         TabIndex        =   6
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label5 
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
         Height          =   348
         Left            =   1824
         TabIndex        =   5
         Top             =   4020
         Width           =   1500
      End
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6204
      Left            =   1860
      Top             =   1332
      Width           =   7932
   End
End
Attribute VB_Name = "frmFAItemCheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  DoEvents
  KillFile "itemchecklist.dat"
  Unload frmFAItemCheckList
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 12 is recommended for this report."
    Call PrintText
  Else
    Exit Sub
  End If
End Sub
Private Sub PrintText()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim MaxLines As Integer
  Dim LineCnt
  Dim ItemCnt
  Dim ReportFile$
  Dim Dash80$, FF$
  Dim Dispose$, Dept$, Index$
  Dim RptHandle As Integer
  Dim Page As Integer, Nextx As Integer
  Dim cnt
  Dim DeptNumber As Double, x As Integer
  Dim TagFlag As Boolean
  Dim NumOfItems As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DeptCodeRec As FADeptCodeType
  Dim DHandle As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim TagIdxCnt As Integer
  Dim DeptDescr$
  Dim DeptCnt As Integer
  Dim FirstFlag As Boolean
  Dim DsplYN$
  Dim ItemTotal As Long
  Dim DeptNumOfItems As Integer
  
  On Error GoTo ERRORSTUFF
  
  FirstFlag = True
  If Check4ValidDept = False Then Exit Sub
  Dispose$ = QPTrim$(fpcmbYN.Text)
  
  TagFlag = False
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    TagFlag = True
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load array with item
    'tag numbers in numeric order
  Next x
  Close TagIdxHandle
  
  ReportFile$ = "FACHK.PRN"     'Report File Name
  Dash80$ = String$(86, "=")
  FF$ = Chr$(12)

  Index$ = QPTrim$(fpcmbOrder.Text)
  
  MaxLines = 60
  LineCnt = 0
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle
  OpenFAItemFile FAHandle

  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  If TagFlag = True Then
   GoSub PrintMasterHeader2
   GoSub TagNumbers
  Else
    GoSub Departments
  End If
  
  GoTo ProcessEnd
  
TagNumbers:
    For cnt = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If LineCnt >= MaxLines Then
        LineCnt = 0
        Print #RptHandle, FF$
        GoSub PrintMasterHeader2
      End If
      'Check For Disposed Of
  
      If Dispose$ = "N" Then 'does user want to include or exclude
      'items already disposed of
      
        'if disposal date is valid and disposal flag = 0 then this item was disposed of
        'before the conversion took place...we want to display items that are in
        'disposal processing (dsplflag = 1) because these items are still in inventory
        'and may need to be counted...dsplflag = 2 means this asset was disposed of
        'after the conversion
        If (FAItemRec.DispDate > 0 And FAItemRec.DsplFlag = 0) Or FAItemRec.DsplFlag = 2 Then GoTo SkipEm2
      Else
        If FAItemRec.DsplFlag = 2 Then
          DsplYN = "Dsp"
        Else
          DsplYN = ""
        End If
      End If
      NumOfItems = NumOfItems + 1 'keep count of all valid items
      Print #RptHandle, DsplYN; Tab(2); FAItemRec.ItemTag;
      Print #RptHandle, Tab(24); Left$(QPTrim$(FAItemRec.IDESC1), 20);
      Print #RptHandle, Tab(47); Left$(QPTrim$(FAItemRec.SERIALNO), 13);
      Print #RptHandle, Tab(61); Left$(QPTrim$(FAItemRec.ITEMLOC), 24);
      Print #RptHandle, Tab(85); "___"
      Print #RptHandle, Tab(2); String$(86, "-")
  
      'SubTotal Here
      If DsplYN = "" Then
        LineCnt = LineCnt + 2
      Else
        LineCnt = LineCnt + 3
      End If
      ItemCnt = ItemCnt + 1
      ItemTotal = ItemTotal + 1
SkipEm2:
      frmFAShowPctComp.ShowPctComp cnt, NumOfFARecs
      If frmFAShowPctComp.Out = True Then
        Close
        frmFAShowPctComp.Out = False
        EnableCloseButton Me.hwnd, True
        Me.cmdExit.Enabled = True
        Me.cmdProcess.Enabled = True
        Unload frmFAShowPctComp
        Exit Sub
      End If
ExitRpt:
  
  Next cnt
  Unload frmFAShowPctComp
  
  Return
  
Departments:
  
  If QPTrim$(fptxtDeptNum.Text) = "" Then
    MsgBox "Please enter a department number"
    fptxtDeptNum.SetFocus
    Close
    Exit Sub
  End If
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb) 'load array of dept nums
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc) 'load array of dept descs
  Next x
  Close DIdxHandle
  
  Dept$ = QPTrim$(fptxtDeptNum.Text) 'if user wants all departments or
  'does he want just one department reported on
  
  'since ALL is a string and cannot be converted into a number if ALL is
  'selected then provisions have to be made for it
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = Val(QPTrim$(DeptNum(x))) Then
        DeptDescr = DeptDesc(x) 'in order to find the correct dept
        'description we must iterate through the department records
        'saved in the DeptDesc() array looking for the corresponding
        'Dept Number
        Exit For 'we found it so jump out of loop
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1))) 'or if all departments are
    'selected then we need the very first dept number and description
    'saved
    DeptDescr = QPTrim(DeptDesc(1))
  End If
  
  Nextx = 1
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) \ Len(TagIdx)
  GoSub PrintMasterHeader
  Do 'this procedure is designed to print out all valid fixed assets
  'in numerical order inside the numerically ordered departments
    For cnt = 1 To TagIdxCnt
      Get FAHandle, TagIdxRecs(cnt), FAItemRec 'retrieve items by number
      If DeptNumber <> FAItemRec.IDEPT Then 'this time thru the loop
      'this item number dept must match the department number for this
      'iteration
        GoTo SkipEm3
      End If
      If Dispose$ = "N" Then 'ok... depts check out but this
      'item is no longer in inventory...if disposal date is valid and
      'disposal amount is more than 0 then this item was disposed of
      'before the conversion and is no longer in inventory...if dsplflag
      '= 2 then this item was disposed of after the conversion and is
      'no longer in inventory...if dsplflag = 1 then this item is in
      'the disposal process but is still in inventory
        If (FAItemRec.DispDate > 0 And FAItemRec.DsplFlag = 0) Or FAItemRec.DsplFlag = 2 Then GoTo SkipEm3
      Else
        If FAItemRec.DsplFlag = 2 Then
          DsplYN = "Dsp"
        Else
          DsplYN = ""
        End If
      End If
      
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader
      End If
      'Check For Disposed Of
      If DeptCnt = 0 Then
        Print #RptHandle,
        Print #RptHandle, Tab(2); "Dept # "; DeptNumber; " "; DeptDescr
        Print #RptHandle, Tab(2); Dash80$
        LineCnt = LineCnt + 3
      End If
      DeptCnt = DeptCnt + 1
      NumOfItems = NumOfItems + 1
      ItemTotal = ItemTotal + 1
      Print #RptHandle, DsplYN; Tab(2); FAItemRec.ItemTag;
      Print #RptHandle, Tab(24); Left$(QPTrim$(FAItemRec.IDESC1), 20);
      Print #RptHandle, Tab(47); Left$(QPTrim$(FAItemRec.SERIALNO), 13);
      Print #RptHandle, Tab(61); Left$(QPTrim$(FAItemRec.ITEMLOC), 24);
      Print #RptHandle, Tab(85); "___"
      Print #RptHandle, Tab(2); String$(86, "-")
  
      'SubTotal Here
      If DsplYN = "" Then
        LineCnt = LineCnt + 2
      Else
        LineCnt = LineCnt + 3
      End If
      ItemCnt = ItemCnt + 1
SkipEm3:

    Next cnt
    
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    
    If QPTrim$(Dept$) <> "ALL" Then Exit Do 'git wagt we wanted with
    'just one loop so exit now
    'so ok to exit
    If DeptCnt > 0 Then
      Print #RptHandle, "Number Of Items: " + CStr(DeptCnt)
      Print #RptHandle, FF$
      If Nextx <> DIdxCnt Then
        GoSub PrintMasterHeader
      End If
    End If
    If Nextx = DIdxCnt Then Exit Do 'been through all departments
    Nextx = Nextx + 1 'otherwise move to next department number
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescr = QPTrim$(DeptDesc(Nextx))
    DeptCnt = 0
  Loop 'loop back to collect data for the next department
  Unload frmFAShowPctComp
  
  Return
  
ProcessEnd:
  If QPTrim$(Dept$) <> "ALL" Then
    GoSub PrintChkEnding
  End If
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  Close         'Close all open files now

  ViewPrint ReportFile$, "Fixed Asset Report", False

  KillFile (ReportFile$)

  Exit Sub

PrintMasterHeader:
       
  Page = Page + 1
  Print #RptHandle, Tab(29); "Asset Check List Report"
  If TagFlag = True Then
    Print #RptHandle, Tab(5); "Listed by Tag Number"
  End If
  
  Print #RptHandle, Tab(2); "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, Tab(2); "Asset Tag Number"; Tab(24); "Description"; ; Tab(47); "Serial Num"; Tab(61); "Location"; Tab(85); "CHK"

  Print #RptHandle, Tab(2); Dash80$
  
  If FirstFlag = True Then
    FirstFlag = False
    LineCnt = 4
  Else
    LineCnt = 5
  End If
  
  Return

PrintMasterHeader2:
  Page = Page + 1
  Print #RptHandle, Tab(29); "Asset Check List Report"
  If TagFlag = True Then
    Print #RptHandle, Tab(2); "Listed by Tag Number"
  Else
    If Dept$ = "ALL" Then
      Print #RptHandle, Tab(2); "Dept # "; DeptNumber; " "; DeptDescr
    Else
      Print #RptHandle, Tab(2); "Dept # "; DeptNumber; " "; DeptDescr
    End If
  End If
  Print #RptHandle, Tab(2); "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, Tab(2); "Asset Tag Number"; Tab(24); "Description"; Tab(47); "Serial Num"; Tab(61); "Location"; Tab(85); "CHK"

  Print #RptHandle, Tab(2); Dash80$
  LineCnt = LineCnt + 4
  Return

PrintChkEnding:
  Print #RptHandle,
  Print #RptHandle, "Number of Items:   "; NumOfItems
  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Unload frmFAShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemCheckList", gstrcProgName, Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
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
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      Call cmdDept_Click
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
      KillFile "itemchecklist.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAItemCheckList.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  'itemchecklist.dat flags this form and tells the department
  'list form (if called) that this was the calling form
  Open "itemchecklist.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fpcmbYN.Text = "N"
  fpcmbYN.AddItem "Y"
  fpcmbYN.AddItem "N"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub fpcmbOrder_Change()
  'disable department number field if user wants report
  'in tag order
  If fpcmbOrder.Text <> "TAG NUMBER" Then
    fptxtDeptNum.Enabled = True
    cmdDept.Enabled = True
  ElseIf fpcmbOrder.Text = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
  End If
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Text = "ALL"
  End If
End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this routine prevents the user from inadvertently changing
  'data in the combo box when scrolling
  If KeyCode = vbKeySpace Then
    fpcmbOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOrder.ListIndex = -1
  End If
  If fpcmbOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
  
End Sub

Private Sub fpcmbYN_KeyDown(KeyCode As Integer, Shift As Integer)
  'this routine prevents the user from inadvertently changing
  'data in the combo box when scrolling
  If KeyCode = vbKeySpace Then
    fpcmbYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYN.ListIndex = -1
  End If
  If fpcmbYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtDeptNum_Change()
  'default this field to ALL if the user deletes the field's value
  If fptxtDeptNum.Text = "" Then
    fptxtDeptNum = "ALL"
  End If
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this routine prevents the user from inadvertently changing
  'data in the combo box when scrolling
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
  'this function compares the department number entered
  'with the valid department numbers on record to validate
  'the entry
  Check4ValidDept = True
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Check4ValidDept = False
    Exit Function
  End If
  
  If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
    Close
    Exit Function
  End If
  
  ThisDept$ = QPTrim$(fptxtDeptNum.Text)
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    If ThisDept$ = QPTrim$(DeptIdx.DeptNumb) Then
      Close
      Exit Function
    End If
  Next x
  Close
  MsgBox "No department number matches this entry. Please try again."
  Check4ValidDept = False
  fptxtDeptNum.SetFocus
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemCheckList", "Check4ValidDept", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Function

Private Sub PrintGraphics()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim ReportFile$
  Dim TagReportFile$
  Dim Dispose$, Dept$, Index$
  Dim RptHandle As Integer
  Dim TagRptHandle As Integer
  Dim Nextx As Integer
  Dim cnt
  Dim DeptNumber As Double, x As Integer
  Dim DeptDescr$
  Dim TagFlag As Boolean
  Dim NumOfItems As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim TagIdxCnt As Integer
  Dim dlm$
  Dim Employer$
  Dim FASHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim ThisDesc As String * 20
  Dim ThisLoc As String * 20
  Dim ThisSerNum As String * 15
  Dim DsplYN As Integer
  Dim ItemTotal As Long
  
  DsplYN = 0
  'Please refer to the PrintText sub for code commenting
  'as it is almost identical to this one
  On Error GoTo ERRORSTUFF
  
  If Check4ValidDept = False Then Exit Sub
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  Employer = FASetUpRec.TownName
  
  dlm$ = "~"
  Dispose$ = QPTrim$(fpcmbYN.Text)
  
  TagFlag = False
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    TagFlag = True
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  ReportFile$ = "FARPTS\FAITEMCHKLIST.RPT"     'Report File Name
  TagReportFile$ = "FARPTS\FATAGITEMCHKLIST.RPT"

  Index$ = QPTrim$(fpcmbOrder.Text)
  
  OpenFAItemFile FAHandle

  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  If TagFlag = True Then
   TagRptHandle = FreeFile
   Open TagReportFile$ For Output As #TagRptHandle
   GoSub TagNumbers
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
    GoSub Departments
  End If
  
  GoTo ProcessEnd
  
TagNumbers:
    For cnt = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If cnt = 1 And Dept$ = "ALL" Then DeptNumber = Val(FAItemRec.IDEPT)
      DsplYN = 0
      If Dispose$ = "N" Then
        If (FAItemRec.DispDate > 0 And FAItemRec.DsplFlag = 0) Or FAItemRec.DsplFlag = 2 Then
          GoTo SkipEm2
        End If
      Else
        If FAItemRec.DsplFlag = 2 Then DsplYN = 1
      End If
      
      NumOfItems = NumOfItems + 1
      ThisDesc = QPTrim$(FAItemRec.IDESC1)
      ThisLoc = QPTrim$(FAItemRec.ITEMLOC)
      ThisSerNum = QPTrim$(FAItemRec.SERIALNO)
      '                        0              1
      Print #TagRptHandle, Employer; dlm; NumOfItems; dlm;
      '                           2
      Print #TagRptHandle, FAItemRec.ItemTag; dlm;
      '                           3
      Print #TagRptHandle, ThisDesc; dlm;
      '                           4           5             6
      Print #TagRptHandle, ThisLoc; dlm; ThisSerNum; dlm; DsplYN
      ItemTotal = ItemTotal + 1
SkipEm2:
      frmFAShowPctComp.ShowPctComp cnt, NumOfFARecs
      If frmFAShowPctComp.Out = True Then
        Close
        frmFAShowPctComp.Out = False
        EnableCloseButton Me.hwnd, True
        Me.cmdExit.Enabled = True
        Me.cmdProcess.Enabled = True
        Unload frmFAShowPctComp
        Exit Sub
      End If
ExitRpt:
  
  Next cnt
  Unload frmFAShowPctComp
  
  Return
  
Departments:
  
  If QPTrim$(fptxtDeptNum.Text) = "" Then
    MsgBox "Please enter a department number"
    fptxtDeptNum.SetFocus
    Close
    Exit Sub
  End If
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  
  'since ALL is a string and cannot be converted into a number if ALL is
  'selected then provisions have to be made for it
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = Val(QPTrim$(DeptNum(x))) Then
        DeptDescr = DeptDesc(x)
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescr = QPTrim$(DeptDesc(1))
  End If
  
  Nextx = 1
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) \ Len(TagIdx)
  
  Do
    For cnt = 1 To TagIdxCnt
      Get TagIdxHandle, cnt, TagIdx
      Get FAHandle, TagIdx.DataRecNum, FAItemRec
      
      If DeptNumber <> FAItemRec.IDEPT Then
        GoTo SkipEm3
      End If
      DsplYN = 0
      If Dispose$ = "N" Then
        If (FAItemRec.DispDate > 0 And FAItemRec.DsplFlag = 0) Or FAItemRec.DsplFlag = 2 Then
          GoTo SkipEm3
        End If
      Else
        If FAItemRec.DsplFlag = 2 Then DsplYN = 1
      End If
  
      NumOfItems = NumOfItems + 1
      ThisDesc = QPTrim$(FAItemRec.IDESC1)
      ThisLoc = QPTrim$(FAItemRec.ITEMLOC)
      ThisSerNum = QPTrim$(FAItemRec.SERIALNO)
      '                     0              1
      Print #RptHandle, Employer; dlm; NumOfItems; dlm;
      '                         2
      Print #RptHandle, FAItemRec.ItemTag; dlm;
      '                         3
      Print #RptHandle, ThisDesc; dlm;
      '                      4            5                6               7              8
      Print #RptHandle, ThisLoc; dlm; DeptNumber; dlm; DeptDescr; dlm; ThisSerNum; dlm; DsplYN
      ItemTotal = ItemTotal + 1
SkipEm3:
  
    Next cnt
    
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    If Nextx = DIdxCnt Then Exit Do
    Nextx = Nextx + 1
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescr = QPTrim$(DeptDesc(Nextx))
    NumOfItems = 0
  Loop
  Unload frmFAShowPctComp
  
  Return
  
ProcessEnd:
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  Close         'Close all open files now
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria.'"
    Close
    Exit Sub
  End If
  
  If TagFlag = True Then
    arFATagItemChkList.Show
  Else
    arFAItemChkList.Show
  End If
  
  frmFALoadReport.Show

  Exit Sub

ERRORSTUFF:
   Unload frmFAShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemCheckList", gstrcProgName, Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Close

End Sub

