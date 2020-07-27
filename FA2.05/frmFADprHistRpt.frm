VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFADprHistRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depreciation History"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmFADprHistRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5895
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10398
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFADprHistRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3555
         TabIndex        =   3
         ToolTipText     =   "Select the preferred method to display this report."
         Top             =   3870
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
         ColDesigner     =   "frmFADprHistRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         ToolTipText     =   "Select the desired order to display this report."
         Top             =   1530
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
         ColDesigner     =   "frmFADprHistRpt.frx":0BDD
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   390
         Left            =   3075
         TabIndex        =   1
         ToolTipText     =   "If DEPARTMENT NUMBER is selected for the Report Order then enter the department to display."
         Top             =   2325
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
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
      Begin EditLib.fpDateTime fptxtDprYear 
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         ToolTipText     =   "Enter the Year to extract W2 information here."
         Top             =   3120
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   656
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
         ButtonStyle     =   2
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
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "2018"
         DateCalcMethod  =   1
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
         DateMax         =   "20350101"
         DateMin         =   "19800101"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "19800101"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
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
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   390
         Left            =   4704
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all departments."
         Top             =   2325
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
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
         ButtonDesigner  =   "frmFADprHistRpt.frx":0ED4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1470
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4680
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
         ButtonDesigner  =   "frmFADprHistRpt.frx":10B4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4440
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   4680
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
         ButtonDesigner  =   "frmFADprHistRpt.frx":1290
      End
      Begin VB.Label Label11 
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
         Left            =   1920
         TabIndex        =   9
         Top             =   2430
         Width           =   930
      End
      Begin VB.Label Label10 
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1590
         Width           =   1830
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation Year:"
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
         Left            =   2190
         TabIndex        =   7
         Top             =   3210
         Width           =   2145
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation History For Year"
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
         Left            =   1830
         TabIndex        =   6
         Top             =   570
         Width           =   4335
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
         Left            =   1875
         TabIndex        =   5
         Top             =   3960
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6150
      Left            =   1780
      Top             =   1185
      Width           =   8050
   End
End
Attribute VB_Name = "frmFADprHistRpt"
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
  KillFile ("dprhistrpt.dat")
  DoEvents
  Unload frmFADprHistRpt
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 17 is recommended for this report."
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim YearChoice$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DprDate$
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim Page As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim DHHandle As Integer
  Dim DprRec As DprHistType
  Dim DprHistCnt As Long
  Dim DDprYr(1) As Double
  Dim DprYr(1) As Double
  Dim ValidCnt As Long
  Dim TempDprRec As DprSortIdxType
  Dim TDHandle As Integer
  Dim DeptDescHeader$, DeptDescription$
  Dim LifeLeft As Double
  Dim DItemCnt As Integer
  Dim FirstFlag As Boolean
  Dim Pct As Integer
  Dim ReversalFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  Pct = 3
  ReversalFlag = False
  FirstFlag = True
  ReportFile$ = "FADPRHISTRPT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  If Check4ValidDept = False Then Exit Sub
  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  YearChoice$ = QPTrim$(fptxtDprYear.Text)
  
'  frmFALoadReport.Show
  DoEvents
  Call CreateDprIdx(YearChoice$, ValidCnt)
'  Unload frmFALoadReport
  
  If ValidCnt = -1 Then
    Close
    MsgBox "No records on file for the selected year."
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim DprIdx(1 To ValidCnt) As Long
  OpenTempDprFile TDHandle
  For x = 1 To ValidCnt
    Get TDHandle, x, TempDprRec
    DprIdx(x) = TempDprRec.DprRecNum 'depreciation files checked out OK
    'so load the array index with records of depreciation records ordered
    'by year and item tag numeric order within that year
  Next x
  Close TDHandle
  
  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text) 'the way the user wants the report displayed
  Open ReportFile$ For Output As #RptHandle
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Unload frmFAShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load asset array with tag
    'numbers in tag numeric order
  Next x
  Close TagIdxHandle
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb) 'need dept #'s loaded into array
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc) 'need dept descriptions loaded into array
  Next x
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt) As Double
  ReDim DTagDYDep(1 To DIdxCnt) As Double
  ReDim DTagDprYr(1 To DIdxCnt) As Double
  ReDim DItemCount(1 To DIdxCnt) As Integer
  
  If Dept$ <> "ALL" Then 'specific dept selected for report
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text)) 'assign dept # on screen
    For x = 1 To DIdxCnt
      If DeptNumber = DeptNum(x) Then 'now get the description for this dept
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1))) 'no specific dept selected so begin
    'report with the first dept
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If
  
  
  OpenFAItemFile FAHandle
  OpenDprHistFile DHHandle
  DprHistCnt = (LOF(DHHandle) / Len(DprRec)) + 1
  If DprHistCnt = 0 Then 'gotta run and post a depreciation first
    MsgBox "No Depreciation History Records on file."
    Close
    Unload frmFAShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  For x = 1 To DprHistCnt
    Get DHHandle, x, DprRec
    If DprRec.DprYear = YearChoice Then
      If DprRec.SoSoftFlag = True Then
        ReversalFlag = True
      End If
      Exit For
    End If
  Next x
  
  GoSub PrintMasterHeader1
  
  TagFlag = False 'used to direct code around dept reporting method
  
GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then 'TagFlag would have been assign TRUE at the
  'end of the first iteration through asset records ...the first time this code is
  'read TagFlag will always be false
    Index = "DEPARTMENT NUMBERS"
    LineCnt = 0
  End If
  
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To ValidCnt
      Get DHHandle, DprIdx(cnt), DprRec
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      'Check For Disposed Date
      DprDate = DprRec.DprYear
      'Check for Acquired Date
      
      If DprDate <> YearChoice Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      YTDDep# = DprRec.DprToDate

      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> DprRec.ThisDept Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      
      If QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" And DItemCnt = 0 Then
        Print #RptHandle,
        Print #RptHandle, "Dept # "; DeptNumber; " "; DeptDescription 'Dept$
        Print #RptHandle, String$(130, "=")
        LineCnt = LineCnt + 3
      End If
      DataFlag = True
      Print #RptHandle, DprRec.ItemTag; Tab(21); DprRec.ThisDesc1;
      Print #RptHandle, Tab(52); DprRec.ThisDept;
      Print #RptHandle, Tab(58); Using("##0", DprRec.Life); "/"; Using("#0", DprRec.LifeLeft);
      Print #RptHandle, Tab(73); Using("###,###,##0.00", CStr(DprRec.OrigCost));
      Print #RptHandle, Tab(88); Using("###,###,##0.00", CStr(DprRec.DprAmt));
      Print #RptHandle, Tab(102); Using("###,###,##0.00", CStr(DprRec.DprToDate));
      Print #RptHandle, Tab(117); Using("###,###,##0.00", CStr(DprRec.BookTotal))
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      DItemCnt = DItemCnt + 1
      LifeLeft = 0
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + DprRec.OrigCost
      BookTotal#(1) = BookTotal#(1) + DprRec.BookTotal
      YDep#(1) = YDep#(1) + YTDDep#
      DprYr#(1) = DprYr#(1) + DprRec.DprAmt
      'collects dept totals
      DOrigCost#(1) = DOrigCost#(1) + DprRec.OrigCost
      DTagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + DprRec.BookTotal
      DTagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      DTagDYDep(Nextx) = DYDep#(1)
      DDprYr#(1) = DDprYr#(1) + DprRec.DprAmt
      DTagDprYr(Nextx) = DDprYr#(1)
      DItemCount(Nextx) = DItemCount(Nextx) + 1
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If
    
    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print
    
    If DataFlag = False Then
      GoTo NoData
    End If
    
  'First Print Subtotals
    Print #RptHandle, String$(130, "-")
    
    Print #RptHandle, "Depreciation for Dept Number: "; DeptNumber; "  "; DeptDescription;
    Print #RptHandle, "  #Items "; DItemCnt;
    Print #RptHandle, Tab(73); Using("###,###,##0.00", CStr(DOrigCost#(1)));
    Print #RptHandle, Tab(88); Using("###,###,##0.00", CStr(DDprYr#(1)));
    Print #RptHandle, Tab(102); Using("###,###,##0.00", CStr(DYDep#(1)));
    Print #RptHandle, Tab(117); Using("###,###,##0.00", CStr(DBookTotal#(1)))
    
    Print #RptHandle, String$(130, "=")
    LineCnt& = LineCnt& + 3
NoData:
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
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DDprYr#(1) = 0
    DItemCnt = 0
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Depreciation Listing", True
  
  KillFile (ReportFile$)
  KillFile (TempDprFileName)
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Depreciation For "; YearChoice
  If FirstFlag = False And QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" Then
    Print #RptHandle, "Dept # "; DeptNumber; " "; DeptDescription 'Dept$
  Else
    Print #RptHandle,
  End If
  If ReversalFlag = True Then
    Print #RptHandle, "*REVERSAL PENDING*"
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(52); "Dept"; Tab(57); "Life/Left"; Tab(76); "Purch Price"; Tab(93); YearChoice$ + " Depr"; Tab(104); "Total Deprec"; Tab(121); "Book Value"
  Print #RptHandle, String$(130, "=")
  LineCnt& = 5
  If FirstFlag = True Then FirstFlag = False
  Return
  
PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Depreciation for Year: "; YearChoice
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(27); "#Items"; Tab(41); "Purchase Price"; Tab(62); YearChoice$ + " Depr"; Tab(75); "Depr To Date"; Tab(93); "Book Value"
  Print #RptHandle, String$(102, "=")
  
  Print #RptHandle, "Total Depreciation "; Tab(29); ItemCnt;
  Print #RptHandle, Tab(41); Using("###,###,##0.00", CStr(OrigCost#(1)));
  Print #RptHandle, Tab(57); Using("###,###,##0.00", CStr(DprYr#(1)));
  Print #RptHandle, Tab(73); Using("###,###,##0.00", CStr(YDep#(1)));
  Print #RptHandle, Tab(89); Using("###,###,##0.00", CStr(BookTotal#(1)))
  
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Depreciation for Year: "; YearChoice
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(5); "Dept"; Tab(19); "Description"; Tab(40); "Items"; Tab(48); "Purchase Price"; Tab(69); YearChoice + " Depr"; Tab(82); "Depr To Date"; Tab(100); "Book Value"
  Print #RptHandle, String$(109, "=")
  LineCnt = 6
  
  
  For x = 1 To DIdxCnt
    Print #RptHandle, Tab(3); Using("##000", DeptNum(x)); Tab(19); DeptDesc(x); Tab(40); Using("####0", CStr(DItemCount(x))); Tab(48); Using("###,###,##0.00", CStr(DTagDOrigCost(x))); Tab(64); Using("###,###,##0.00", CStr(DTagDprYr(x))); Tab(80); Using("###,###,##0.00", CStr(DTagDYDep(x))); Tab(96); Using("###,###,##0.00", CStr(DTagDBookTotal(x)))
    LineCnt = LineCnt + 1
    
    If LineCnt& >= MaxLines And x <> DIdxCnt Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Depreciation Year "; YearChoice
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, Tab(5); "Dept"; Tab(19); "Description"; Tab(40); "Items"; Tab(46); "Purchase Price"; Tab(62); YearChoice + " Depr"; Tab(75); "Depr To Date"; Tab(93); "Book Value"
      Print #RptHandle, String$(102, "=")
      LineCnt = LineCnt + 6
    End If
  Next x
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistRpt", "PrintText", Erl)
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
    Unload Me
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
'    'Me.Visible = False
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
      KillFile ("dprhistrpt.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADprHistRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  'Disable dept num field if tag number is seletced
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then 'default to TAG NUMBER
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    fptxtDeptNum.Enabled = True
    cmdDept.Enabled = True
  End If

End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub LoadMe()
  Dim One As Integer
  Dim FileHandle As Integer
  One = 1
  FileHandle = FreeFile
  Open "dprhistrpt.dat" For Output As FileHandle Len = 2 'tells
  'dept list that the request for a dept number from the list comes
  'from here
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  
End Sub

Private Sub fpcomboPrintOpt_Change()
  'default to Graphical
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fptxtDeptNum_Change()
  'default to ALL
  If QPTrim$(fptxtDeptNum.Text) = "" Then
    fptxtDeptNum = "ALL"
  End If
End Sub

Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
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
  
  MsgBox "No department number matches this entry. Please try again."
  Check4ValidDept = False
  fptxtDeptNum.SetFocus
  Close
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistRpt", "Check4ValidDept", Erl)
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
    Unload Me

End Function

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub PrintGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim ItemCnt&
  Dim Dept$
  Dim YearChoice$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DprDate$
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim DHHandle As Integer
  Dim DprRec As DprHistType
  Dim DprHistCnt As Long
  Dim DDprYr(1) As Double
  Dim DprYr(1) As Double
  Dim ValidCnt As Long
  Dim TempDprRec As DprSortIdxType
  Dim TDHandle As Integer
  Dim DeptDescHeader$, DeptDescription$
  Dim TagReportFile$
  Dim TagRptHandle As Integer
  Dim SubReportFile$
  Dim SubTagHandle As Integer
  Dim Employer$
  Dim FASHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim dlm$, DItemCnt As Long
  Dim DeptFlag As Boolean
  Dim Pct As Integer
  
  On Error GoTo ERRORSTUFF
  
  Pct = 4
  DeptFlag = False
  dlm = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer = FASetUpRec.TownName

  ReportFile$ = "FARPTS\FADEPRHIST.RPT"  'Report File Name
  TagReportFile = "FARPTS\FATAGDEPRHIST.RPT"
  SubReportFile = "FARPTS\FASUBDEPRHIST.RPT"
  If Check4ValidDept = False Then Exit Sub
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  YearChoice$ = QPTrim$(fptxtDprYear.Text)
  Index$ = QPTrim$(fpcmbOrder.Text)
  DoEvents
  Call CreateDprIdx(YearChoice$, ValidCnt)
  
  If ValidCnt = -1 Then
    Close
    MsgBox "No records on file for the selected year."
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
    Exit Sub
  End If
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim DprIdx(1 To ValidCnt) As Long
  OpenTempDprFile TDHandle
  For x = 1 To ValidCnt
    Get TDHandle, x, TempDprRec
    DprIdx(x) = TempDprRec.DprRecNum
  Next x
  Close TDHandle
  
  If Index = "TAG NUMBER" Then
    TagRptHandle = FreeFile
    Open TagReportFile$ For Output As #TagRptHandle
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
    DeptFlag = True
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Unload frmFAShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
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
  
  ReDim DTagDOrigCost(1 To DIdxCnt) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt) As Double
  ReDim DTagDYDep(1 To DIdxCnt) As Double
  ReDim DTagDprYr(1 To DIdxCnt) As Double
  ReDim DTagItemCnt(1 To DIdxCnt) As Long
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt ' + 1
      If DeptNumber = DeptNum(x) Then
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If
  
  OpenFAItemFile FAHandle
  OpenDprHistFile DHHandle
  DprHistCnt = (LOF(DHHandle) / Len(DprRec)) + 1
  If DprHistCnt = 0 Then
    MsgBox "No Depreciation History Records on file."
    Close
    Unload frmFAShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  TagFlag = False
  
GetTagTotals:
  Nextx = 1
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
  End If
  
  Do
    DataFlag = False
    For cnt& = 1 To ValidCnt
      Get DHHandle, DprIdx(cnt), DprRec
      'Check For Disposed Date
      DprDate = Mid(DprRec.DprYear, 1, 4)
      'Check for Acquired Date
      
      If DprDate <> YearChoice Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      YTDDep# = DprRec.DprToDate

      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1
      ElseIf DeptNumber <> DprRec.ThisDept Then
        GoTo SkipEm1
      End If
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If Index$ = "TAG NUMBER" Then
        DataFlag = True
        '                        0                1                        2
        Print #TagRptHandle, Employer; dlm; DprRec.ItemTag; dlm; DprRec.ThisDesc1; dlm;
        '                            3                  4                    5
        Print #TagRptHandle, DprRec.ThisDept; dlm; DprRec.Life; dlm; DprRec.LifeLeft; dlm;
        '                           6                    7                     8
        Print #TagRptHandle, DprRec.OrigCost; dlm; DprRec.DprAmt; dlm; DprRec.DprToDate; dlm;
        '                            9                  10                  11                     12
        Print #TagRptHandle, DprRec.BookTotal; dlm; YearChoice$; dlm; FAItemRec.DEPYN; dlm; DprRec.SoSoftFlag
      Else
        '                        0                1                        2
        Print #RptHandle, Employer; dlm; DprRec.ItemTag; dlm; DprRec.ThisDesc1; dlm;
        '                            3                  4                    5
        Print #RptHandle, DprRec.ThisDept; dlm; DprRec.Life; dlm; DprRec.LifeLeft; dlm;
        '                           6                    7                     8
        Print #RptHandle, DprRec.OrigCost; dlm; DprRec.DprAmt; dlm; DprRec.DprToDate; dlm;
        '                            9                  10                11
        Print #RptHandle, DprRec.BookTotal; dlm; YearChoice$; dlm; FAItemRec.DEPYN; dlm;
        '                      12                  13                 14
        Print #RptHandle, DeptNumber; dlm; DeptDescription; dlm; DOrigCost#(1); dlm;
        '                      15              16         17
        Print #RptHandle, DDprYr#(1); dlm; DYDep#(1); dlm; DBookTotal#(1); dlm; DprRec.SoSoftFlag
      End If


TagOnly2:
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      DItemCnt = DItemCnt + 1

      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + DprRec.OrigCost
      BookTotal#(1) = BookTotal#(1) + DprRec.BookTotal
      YDep#(1) = YDep#(1) + YTDDep#
      DprYr#(1) = DprYr#(1) + DprRec.DprAmt
      'collects dept totals
      DOrigCost#(1) = DOrigCost#(1) + DprRec.OrigCost
      DTagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + DprRec.BookTotal
      DTagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      DTagDYDep(Nextx) = DYDep#(1)
      DDprYr#(1) = DDprYr#(1) + DprRec.DprAmt
      DTagDprYr(Nextx) = DDprYr#(1)
      DTagItemCnt(Nextx) = DItemCnt

SkipEm1:

    Next cnt&
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If
    
    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print
    
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt + 1
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
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DDprYr#(1) = 0
    DItemCnt = 0
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  Close         'Close all open files now
  
  If DeptFlag = False Then
    arFATagDprHistRpt.Show
  Else
    arFADprHistRpt.Show
  End If
  
  frmFALoadReport.Show
  
  Exit Sub
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  SubTagHandle = FreeFile
  Open SubReportFile For Output As SubTagHandle
  For x = 1 To DIdxCnt
    '                        0                1                  2                      3                  4                    5                    6                  7
    Print #SubTagHandle, DeptNum(x); dlm; DeptDesc(x); dlm; DTagDOrigCost(x); dlm; DTagDprYr(x); dlm; DTagDYDep(x); dlm; DTagDBookTotal(x); dlm; YearChoice; dlm; DTagItemCnt(x)
  Next x
  Close SubTagHandle
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADprHistRpt", "PrintGraphics", Erl)
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
    Unload Me

End Sub

