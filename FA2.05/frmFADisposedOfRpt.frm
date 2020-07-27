VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFADisposedOfRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items Disposed Of Report"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmFADisposedOfRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6204
      Left            =   1944
      TabIndex        =   7
      Top             =   1332
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   10943
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFADisposedOfRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   384
         Left            =   3516
         TabIndex        =   4
         Top             =   4032
         Width           =   2364
         _Version        =   196608
         _ExtentX        =   4170
         _ExtentY        =   677
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
         ColDesigner     =   "frmFADisposedOfRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   384
         Left            =   3216
         TabIndex        =   0
         ToolTipText     =   "Select the order in which you wish this report  to be displayed."
         Top             =   1680
         Width           =   3228
         _Version        =   196608
         _ExtentX        =   5694
         _ExtentY        =   677
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
         ColDesigner     =   "frmFADisposedOfRpt.frx":0BA5
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "F8 &Dept List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   4416
         TabIndex        =   2
         ToolTipText     =   "Click this button to bring up a complete list of all departments."
         Top             =   2448
         Width           =   1356
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC &Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   684
         Left            =   1584
         TabIndex        =   5
         Top             =   4944
         Width           =   1884
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "F10 &Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   684
         Left            =   4464
         TabIndex        =   6
         Top             =   4944
         Width           =   1884
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   1
         ToolTipText     =   "If DEPARTMENT is selected for the Report Order then select the department you wish to display."
         Top             =   2448
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
      Begin EditLib.fpDateTime fptxtDispYear 
         Height          =   372
         Left            =   4608
         TabIndex        =   3
         ToolTipText     =   "Enter the year "
         Top             =   3264
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
         Text            =   "2003"
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
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
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
         Left            =   1440
         TabIndex        =   12
         Top             =   1776
         Width           =   1548
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
         TabIndex        =   11
         Top             =   2544
         Width           =   924
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Items Disposal Year:"
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
         TabIndex        =   10
         Top             =   3312
         Width           =   2364
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   480
         Width           =   4908
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Items Disposed Of Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Left            =   1584
         TabIndex        =   9
         Top             =   624
         Width           =   4812
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   8
         Top             =   4116
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6396
      Left            =   1860
      Top             =   1236
      Width           =   7932
   End
End
Attribute VB_Name = "frmFADisposedOfRpt"
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
  DoEvents
  KillFile "disposedofrpt.dat"
  Unload frmFADisposedOfRpt

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
Private Sub PrintGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim TDsplAmt As Double
  Dim DsplAmt As Double
  Dim ReportFile$
  Dim ItemCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
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
  Dim DisposalDate As Integer
  Dim TagReportFile$
  Dim TagHandle As Integer
  Dim dlm$
  Dim TagSubHandle As Integer
  Dim TagSubReportFile$
  Dim FASetUpRec As FASetupRecType
  Dim Employer$
  Dim YDispPrice#(2), DYDispPrice#(2)
  Dim DispPrice#
  Dim DCnt As Integer
  Dim Method$
  Dim ActiveX As Long
  Dim HoldRec As Long
  Dim HoldDate As Integer
  Dim ThisDRec As Long
  Dim BigNum As Long
  Dim SmallNum As Long
  Dim HoldTag$
  
  If Check4ValidDept = False Then Exit Sub

  OpenFASetUpFile FAHandle
  Get FAHandle, 1, FASetUpRec
  Close FAHandle
  Employer = FASetUpRec.TownName
  
  dlm$ = "~"
  ReportFile$ = "FARPTS\FADISPOSEDOF.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FADISPOSEDOFTAG.RPT"
  TagSubReportFile$ = "FARPTS\FADISPOSEOFSUB.RPT"
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num("01/01/" & fptxtDispYear)
  EDate = Date2Num("12/31/" & fptxtDispYear)
  
  Index$ = QPTrim$(fpcmbOrder.Text)
  If QPTrim$(Index$) = "DEPARTMENT NUMBER" Then
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
  Else
    TagHandle = FreeFile
    Open TagReportFile$ For Output As #TagHandle
    TagSubHandle = FreeFile
    Open TagSubReportFile$ For Output As #TagSubHandle
  End If
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Close
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
  ReDim DeptArr(1 To DIdxCnt + 1) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
  Next x
  DeptArr(x) = ""
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim DTagDYDep(1 To DIdxCnt + 1) As Double
  ReDim DTagDYDispPrice(1 To DIdxCnt + 1) As Double
  ReDim DTagDCnt(1 To DIdxCnt + 1) As Integer
  ReDim DDsplAmt(1 To DIdxCnt + 1) As Double
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
  End If
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  'start date sort
  ReDim ItemsForDisposal(1 To 1) As Long
  ReDim DatesForDisposal(1 To 1) As Integer
  For x = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If FAItemRec.DispDate >= BDate And DisposeDate <= EDate And FAItemRec.DsplFlag = 2 Then
      ActiveX = ActiveX + 1
      ReDim Preserve ItemsForDisposal(1 To ActiveX) As Long
      ReDim Preserve DatesForDisposal(1 To ActiveX) As Integer
      ItemsForDisposal(ActiveX) = TagIdxRecs(x)
      DatesForDisposal(ActiveX) = FAItemRec.DispDate
    End If
  Next x
  
  If ActiveX = 0 Then
    MsgBox "No fixed assets could be found that were disposed of within the year entered."
    fptxtDispYear.SetFocus
    Close
    Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To ActiveX
    If DatesForDisposal(x) > BigNum Then
      BigNum = DatesForDisposal(x)
    End If
  Next x
  
  Nextx = 1
  SmallNum = BigNum + 1
  Do
    For x = Nextx To ActiveX ' - 1
      If DatesForDisposal(x) < SmallNum Then
        SmallNum = DatesForDisposal(x)
        ThisDRec = x
      End If
    Next x
    HoldRec = ItemsForDisposal(ThisDRec)
    HoldDate = DatesForDisposal(ThisDRec)
    ItemsForDisposal(ThisDRec) = ItemsForDisposal(Nextx)
    DatesForDisposal(ThisDRec) = DatesForDisposal(Nextx)
    ItemsForDisposal(Nextx) = HoldRec
    DatesForDisposal(Nextx) = HoldDate
    If Nextx = ActiveX Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  'end date sort
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False

GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
  End If
  Nextx = 1
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To ActiveX ' - 1   'NumOfFARecs
      Get FAHandle, ItemsForDisposal(cnt), FAItemRec
'      Get FAHandle, TagIdxRecs(Cnt), FAItemRec
      'Check For Disposed Date
      DisposeDate = FAItemRec.DispDate
      'Check for Acquired Date
      AcquireDate = FAItemRec.AQURDATE
      
      If DisposeDate < BDate Or DisposeDate > EDate Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      If FAItemRec.DsplFlag < 2 Then GoTo SkipEm1
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If QPTrim$(FAItemRec.DsplMethod) = "Salvage" Then
        Method$ = "SLV"
      ElseIf QPTrim$(FAItemRec.DsplMethod) = "Sold" Then
        Method$ = "SLD"
      ElseIf QPTrim$(FAItemRec.DsplMethod) = "Auction" Then
        Method$ = "AUC"
      Else
        Method$ = "OTH"
      End If
      
      DispPrice# = FAItemRec.DisposAmt
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      
      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
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
      DataFlag = True
      
      If QPTrim$(Index) = "TAG NUMBER" Then
        '                   0               1                 2
        Print #TagHandle, Dept$; dlm; fptxtDispYear; dlm; Employer$; dlm;
        '                          3                          4
        Print #TagHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
        '                        5                     6                       7
        Print #TagHandle, FAItemRec.IDEPT; dlm; Method$; dlm; FAItemRec.ORGCOST; dlm;
'          '                    8
        Print #TagHandle, FAItemRec.DEP2DATE; dlm;
        '                                    9                                      10                              11
        Print #TagHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.DispDate); dlm; FAItemRec.DisposAmt
      Else
        '                   0               1                2
        Print #RptHandle, Dept$; dlm; fptxtDispYear; dlm; Employer$; dlm;
        '                         3                             4
        Print #RptHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
        '                        5                    6                       7
        Print #RptHandle, FAItemRec.IDEPT; dlm; Method; dlm; FAItemRec.ORGCOST; dlm;
          '                    8
        Print #RptHandle, FAItemRec.DEP2DATE; dlm;
        '                                   9                                       10
        Print #RptHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.DispDate); dlm;
      End If
      
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      DCnt = DCnt + 1
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
      BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
      YDep#(1) = YDep#(1) + YTDDep#
      YDispPrice#(1) = YDispPrice#(1) + DYDispPrice#(1)
      DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
      DTagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
      DTagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      DTagDYDep(Nextx) = DYDep#(1)
      DYDispPrice#(1) = DYDispPrice#(1) + DispPrice#
      DTagDYDispPrice(Nextx) = DTagDYDispPrice(Nextx) + FAItemRec.DisposAmt
      DTagDCnt(Nextx) = DCnt
      If QPTrim$(Index$) = "DEPARTMENT NUMBER" Then
        '                     11               12                 13
        Print #RptHandle, DeptNumber; dlm; DOrigCost#(1); dlm; DYDep#(1); dlm;
        '                     14                   15                16              17
        Print #RptHandle, DBookTotal#(1); dlm; OrigCost#(1); dlm; YDep#(1); dlm; BookTotal#(1); dlm;
        '                     18                    19
        Print #RptHandle, DYDispPrice#(1); dlm; DispPrice#
      End If
      
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
'      Print #RptHandle, "NO DELETIONS FOR DEPARTMENT "; DeptNumber
'      Print #RptHandle, String$(123, "=")
'      LineCnt = LineCnt + 1
      GoTo NoData
    End If
    
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
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt + 1 Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DCnt = 0
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
  Close RptHandle
  If TagFlag = False Then
    arFADisposedOfRpt.Show
  Else
    arFADisposedOFTagOnly.Show
  End If
  
  frmFALoadReport.Show
  
  Exit Sub
  
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  For x = 1 To DIdxCnt + 1
    If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
    '                        0                    1                   2                    3                          4
    Print #TagSubHandle, DeptArr(x); dlm; DTagDOrigCost(x); dlm; DTagDYDep(x); dlm; DTagDBookTotal(x); dlm; DTagDYDispPrice(x); dlm; DTagDCnt(x)
  Next x
  Return

End Sub
Private Sub PrintText()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim TDsplAmt As Double
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
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
  Dim DisposalDate As Integer
  Dim DsplAmt As Double
  Dim Method$
  Dim ActiveX As Long
  Dim HoldRec As Long
  Dim HoldDate As Integer
  Dim ThisDRec As Long
  Dim BigNum As Long
  Dim SmallNum As Long
  Dim HoldTag$
  
  If Check4ValidDept = False Then Exit Sub
  
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num("01/01/" & fptxtDispYear)
  EDate = Date2Num("12/31/" & fptxtDispYear)
  
  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text)
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintMasterHeader1
  
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
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptArr(1 To DIdxCnt + 1) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
  Next x
  DeptArr(x) = ""
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim DTagDYDep(1 To DIdxCnt + 1) As Double
  ReDim DDsplAmt(1 To DIdxCnt + 1) As Double
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
  End If
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  'start date sort
  ReDim ItemsForDisposal(1 To 1) As Long
  ReDim DatesForDisposal(1 To 1) As Integer
  For x = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If FAItemRec.DispDate >= BDate And DisposeDate <= EDate And FAItemRec.DsplFlag = 2 Then
      ActiveX = ActiveX + 1
      ReDim Preserve ItemsForDisposal(1 To ActiveX) As Long
      ReDim Preserve DatesForDisposal(1 To ActiveX) As Integer
      ItemsForDisposal(ActiveX) = TagIdxRecs(x)
      DatesForDisposal(ActiveX) = FAItemRec.DispDate
    End If
  Next x
  
  If ActiveX = 0 Then
    MsgBox "No fixed assets could be found that were disposed of within the year entered."
    fptxtDispYear.SetFocus
    Close
    Exit Sub
  End If
  
  BigNum = 0
  For x = 1 To ActiveX
    If DatesForDisposal(x) > BigNum Then
      BigNum = DatesForDisposal(x)
    End If
  Next x
  
  Nextx = 1
  SmallNum = BigNum + 1
  Do
    For x = Nextx To ActiveX ' - 1
      If DatesForDisposal(x) < SmallNum Then
        SmallNum = DatesForDisposal(x)
        ThisDRec = x
      End If
    Next x
    HoldRec = ItemsForDisposal(ThisDRec)
    HoldDate = DatesForDisposal(ThisDRec)
    ItemsForDisposal(ThisDRec) = ItemsForDisposal(Nextx)
    DatesForDisposal(ThisDRec) = DatesForDisposal(Nextx)
    ItemsForDisposal(Nextx) = HoldRec
    DatesForDisposal(Nextx) = HoldDate
    If Nextx = ActiveX Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  'end date sort
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False

GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
    LineCnt = 0
  End If
  
  
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To ActiveX ' - 1   'NumOfFARecs
      Get FAHandle, ItemsForDisposal(cnt), FAItemRec
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      'Check For Disposed Date
      DisposeDate = FAItemRec.DispDate
      'Check for Acquired Date
      AcquireDate = FAItemRec.AQURDATE
      
      If DisposeDate < BDate Or DisposeDate > EDate Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      If FAItemRec.DsplFlag < 2 Then GoTo SkipEm1
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
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
      
      DataFlag = True
      
      Print #RptHandle, FAItemRec.ItemTag; Tab(21); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(50); FAItemRec.IDEPT;
      If QPTrim$(FAItemRec.DsplMethod) = "Salvage" Then
        Method$ = "SLV"
      ElseIf QPTrim$(FAItemRec.DsplMethod) = "Sold" Then
        Method$ = "SLD"
      ElseIf QPTrim$(FAItemRec.DsplMethod) = "Auction" Then
        Method$ = "AUC"
      Else
        Method$ = "OTH"
      End If
      Print #RptHandle, Tab(57); Method;
      Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
      Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL));
      Print #RptHandle, Tab(94); Using("###,###,##0.00", CStr(FAItemRec.DisposAmt));
      Print #RptHandle, Tab(116); MakeRegDate(FAItemRec.DispDate)
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
      BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
      DsplAmt = DsplAmt + FAItemRec.DisposAmt
      YDep#(1) = YDep#(1) + YTDDep#
      DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
      DTagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
      DTagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      DTagDYDep(Nextx) = DYDep#(1)
      DDsplAmt(Nextx) = DDsplAmt(Nextx) + FAItemRec.DisposAmt
      TDsplAmt = TDsplAmt + FAItemRec.DisposAmt
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
'      Print #RptHandle, "NO DELETIONS FOR DEPARTMENT "; DeptNumber
'      Print #RptHandle, String$(123, "=")
'      LineCnt = LineCnt + 1
      GoTo NoData
    End If
    
  'First Print Subtotals
    Print #RptHandle,
    Print #RptHandle, "Totals for Dept Number: "; DeptNumber;
    Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(DOrigCost#(1)));
    Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(DBookTotal#(1)));
    Print #RptHandle, Tab(94); Using("###,###,##0.00", CStr(DsplAmt))
    
    Print #RptHandle, String$(125, "=")
    LineCnt& = LineCnt& + 4
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
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt + 1 Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
'    DOrigCost#(2) = 0
'    DBookTotal#(2) = 0
'    DYDep#(2) = 0
    DsplAmt = 0
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
  
  Close
  ViewPrint ReportFile$, "Master Asset Listing", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Disposed Items"
  Print #RptHandle, "Dept # "; Dept$
  Print #RptHandle, "Items Disposed From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
'  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(50); "Dept"; Tab(55); "Method"; Tab(63); "Original Cost"; Tab(81); "Book Value"; Tab(94); "Disposal Price"; Tab(113); "Disposal Date"
  Print #RptHandle, String$(125, "=")
  LineCnt& = 7
  Return
  
PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(18); "Total Purchase Price"; Tab(45); "Total Book Value"; Tab(63); "Total Disposal Price "
  Print #RptHandle, String$(83, "=")
  
  Print #RptHandle, "All Deleted Items ";
  Print #RptHandle, Tab(24); Using("###,###,##0.00", CStr(OrigCost#(1)));
  Print #RptHandle, Tab(47); Using("###,###,##0.00", CStr(BookTotal#(1)));
  Print #RptHandle, Tab(69); Using("###,###,##0.00", CStr(TDsplAmt))
  
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "Dept Number"; Tab(18); "Total Purchase Price"; Tab(45); "Total Book Value"; Tab(63); "Total Disposal Price"
  Print #RptHandle, String$(83, "=")
  LineCnt = 5
  
  
  For x = 1 To DIdxCnt + 1
    If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
    Print #RptHandle, Tab(4); Using$("###0", DeptArr(x)); Tab(24); Using("###,###,##0.00", CStr(DTagDOrigCost(x))); Tab(47); Using("###,###,##0.00", CStr(DTagDBookTotal(x))); Tab(69); Using("###,###,##0.00", CStr(DDsplAmt(x)))
    LineCnt = LineCnt + 1
    
    If LineCnt& >= MaxLines And x <> DIdxCnt + 1 Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, "Dept Number"; Tab(18); "Total Purchase Price"; Tab(45); "Total Book Value"; Tab(63); "Total Disposal Price"
      Print #RptHandle, String$(83, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  Return

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call Loadme
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
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
      KillFile "disposedofrpt.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADisposedOfRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    fptxtDeptNum.Enabled = True
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

Private Sub Loadme()
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  Open "disposedofrpt.dat" For Output As FileHandle Len = 2 'tells
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
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fptxtDeptNum_DblClick(Button As Integer)
'  Dim This$
'  This$ = Clipboard.GetText
'  If This$ = "" Then Exit Sub
'  fptxtDeptNum = Clipboard.GetText
'  Clipboard.Clear

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
  
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

