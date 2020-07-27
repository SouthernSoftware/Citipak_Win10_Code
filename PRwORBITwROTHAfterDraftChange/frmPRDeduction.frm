VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPRDeduction 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PayRoll Deduction Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmPRDeduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7215
      Left            =   2160
      TabIndex        =   3
      Top             =   810
      Width           =   7305
      _Version        =   196609
      _ExtentX        =   12885
      _ExtentY        =   12726
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Picture         =   "frmPRDeduction.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3555
         TabIndex        =   2
         Top             =   5520
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
         ColDesigner     =   "frmPRDeduction.frx":08E6
      End
      Begin VB.ListBox fplistDedNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   3480
         MultiSelect     =   1  'Simple
         TabIndex        =   10
         Top             =   1440
         Width           =   2895
      End
      Begin EditLib.fpDateTime fptxtStartDate 
         Height          =   465
         Left            =   1080
         TabIndex        =   0
         Top             =   1920
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
         _ExtentY        =   820
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
         Text            =   "11/20/2002"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEndDate 
         Height          =   465
         Left            =   1080
         TabIndex        =   1
         Top             =   2835
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
         _ExtentY        =   820
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
         Text            =   "11/20/2002"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4230
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the desired employee deduction report.."
         Top             =   6195
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
         ButtonDesigner  =   "frmPRDeduction.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1290
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   6195
         Width           =   1905
         _Version        =   131072
         _ExtentX        =   3360
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
         ButtonDesigner  =   "frmPRDeduction.frx":0DBC
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   570
         Left            =   840
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4320
         Width           =   2370
         _Version        =   131072
         _ExtentX        =   4180
         _ExtentY        =   1005
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
         ButtonDesigner  =   "frmPRDeduction.frx":0F9A
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdAll 
         Height          =   570
         Left            =   840
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2370
         _Version        =   131072
         _ExtentX        =   4180
         _ExtentY        =   1005
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
         ButtonDesigner  =   "frmPRDeduction.frx":1181
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
         Left            =   1440
         TabIndex        =   7
         Top             =   5610
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1290
         Top             =   315
         Width           =   4815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D0D0D0&
         Caption         =   "End Date:"
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
         TabIndex        =   6
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Deduction Report"
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
         Height          =   450
         Left            =   1440
         TabIndex        =   4
         Top             =   510
         Width           =   4530
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   7515
      Left            =   2010
      Top             =   675
      Width           =   7650
   End
End
Attribute VB_Name = "frmPRDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim AllFlag As Boolean
Dim GDedCount As Integer

Private Sub cmdAll_Click()
  Dim x As Integer
  
  For x = 0 To GDedCount - 1
    fplistDedNo.Selected(x) = True
  Next x
End Sub

Private Sub cmdClear_Click()
  Dim x As Integer
  
  For x = 0 To GDedCount - 1
   fplistDedNo.Selected(x) = False
  Next x
End Sub

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmPRDeduction
End Sub

Private Sub cmdProcess_Click()
  If fplistDedNo.SelCount = 0 Then
    MsgBox "Please make a selection from the deduction list."
    Exit Sub
  End If
  
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
    Case vbKeyReturn:
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdEscape_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF4:
      Call cmdAll_Click
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF2:
      Call cmdClear_Click
      SendKeys "%L"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadThisForm
  Me.HelpContextID = hlpPayrollDeductions
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Sub PrintGraphics()

  Dim IdxRecLen As Integer
  Dim FormName As String, x As Integer, y As Integer
  Dim cnt As Integer, DAmt#
  Dim LowDate As Long
  Dim HiDate As Long, TransRecNum&
  Dim EmpRecSize As Long, TempDed$, NumOfRecs As Long
  Dim DNum As Integer
  Dim DedCodeFile As DedCodeRecType, RecNo As Integer
  Dim Image1$, Image2$, RptTitle$
  Dim RptName$, RptDate$, UsingThisOne As Boolean
  Dim SubRptName$, SubHandle As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim TransHistFileName As TransRecType
  Dim TotalDAmt#, UFileHandle As Integer
  Dim DedCodeFileHandle As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim IdxFileName As NameSortIdxType
  Dim Emp2Rec As EmpData2Type
  Dim dlm$
  Dim NumOfDeds As Integer
  Dim GTotalDAmt#, ThisDed As DedCodeRecType
  Dim SSN$, TotActive As Integer
  Dim Sub2RptName$
  Dim Sub2Handle As Integer
  Dim RptCnt As Integer
  Dim Nextx As Integer
  Dim ThisNumOfDeds As Integer
  
  On Error GoTo ERRORSTUFF
  RptCnt = 0
  dlm = "~"
  OpenEmpIdxLNameFile EmpIdxLNameHandle

  NumOfRecs = LOF(EmpIdxLNameHandle) \ 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  Close EmpIdxLNameHandle
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "ERROR: The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  If CheckValDate(fptxtEndDate.Text) = False Then
    MsgBox "ERROR: The End Date is not valid"
    fptxtEndDate.SetFocus
    Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
    MsgBox "ERROR:The Start Date is later than the End Date"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If fplistDedNo.Text = "" Then
    MsgBox "ERROR: Please make a selection in the Deduction Number field"
    fplistDedNo.SetFocus
    Exit Sub
  End If
  If fptxtStartDate.Text = "" Then
    MsgBox "ERROR: Please enter a Starting Date"
    fptxtStartDate.SetFocus
    Exit Sub
  End If

  If fptxtEndDate.Text = "" Then
    MsgBox "ERROR: Please enter an Ending Date"
    fptxtEndDate.SetFocus
    Exit Sub
  End If
  
  OpenDedCodeFile DedCodeFileHandle
  NumOfDeds = LOF(DedCodeFileHandle) / Len(ThisDed)
  ReDim DedCodes(1 To NumOfDeds) As DedCodeRecType
  For cnt = 1 To NumOfDeds '50 'load up DedCodes array
    Get DedCodeFileHandle, cnt, DedCodes(cnt)
  Next
  
  Nextx = 0
  ReDim ThisSel(1 To fplistDedNo.SelCount) As Integer
  For x = 0 To NumOfDeds - 1
    If fplistDedNo.Selected(x) = True Then
      Nextx = Nextx + 1
      ThisSel(Nextx) = x + 1
    End If
  Next x

  Image1$ = "$#,##0.00"
  Image2$ = "$###,##0.00"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim DedRpt(1) As DedRptType

  OpenUnitFile UFileHandle
  Get UFileHandle, 1, Unit(1)
  IdxRecLen = 2
  
  OpenEmpIdxLNameFile EmpIdxLNameHandle

  NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen

  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  
  RptName$ = "PRRPTS\DEDUCTG" & DNum & ".RPT"
  RptDate$ = "Start Date: " + MakeRegDate(LowDate) + "  Ending Date: " + MakeRegDate(HiDate) ''Num2Date
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  
  RptName$ = "PRRPTS\DEDUCALL.RPT"
  SubRptName$ = "PRRPTS\SUBDEDALL.RPT"
  Sub2RptName$ = "PRRPTS\SUB2DEDALL.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  SubHandle = FreeFile
  
  Open SubRptName$ For Output As SubHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  ReDim DedEmpAmt(1 To NumOfDeds, 1 To NumOfRecs) As Double
  ReDim DedTotAmt(1 To NumOfDeds) As Double
  ReDim DedEmpCnt(1 To NumOfDeds) As Integer
  frmLoadingRpt.Show
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    'If this employee has not been terminated
    'If there are no transactions for this employee then skip 'em
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm2All
    End If
      TransRecNum& = CLng(Emp2Rec.LastTransRec)
      Do
        Get THandle, TransRecNum&, TransHRec(1)
        Select Case TransHRec(1).CheckDate
        'if this transaction falls within the date parameters
        Case LowDate To HiDate
        'if this employee has a value for this transaction for this deduction
        'then assign that value to DAmt and flag this as being used
          For y = 1 To fplistDedNo.SelCount
            If TransHRec(1).DAmt(ThisSel(y)) <> 0 Then
              DedEmpAmt(ThisSel(y), CLng(IdxBuff(RecNo))) = DedEmpAmt(ThisSel(y), CLng(IdxBuff(RecNo))) + TransHRec(1).DAmt(ThisSel(y))
              DedTotAmt(ThisSel(y)) = OldRound(DedTotAmt(ThisSel(y)) + TransHRec(1).DAmt(ThisSel(y)))
            End If
          Next y
        Case Else
        End Select
        If TransHRec(1).PrevTransRec <= 0 Then
          Exit Do
        Else
          TransRecNum& = CLng(TransHRec(1).PrevTransRec)
        End If
      Loop
SkipEm2All:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Unload frmLoadingRpt
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  
  Nextx = 1
  For x = 1 To NumOfDeds
    If Nextx = fplistDedNo.SelCount + 1 Then Exit For
    If ThisSel(Nextx) = x Then
      Get DedCodeFileHandle, x, ThisDed
      Nextx = Nextx + 1
    Else
      GoTo SkipDed
    End If
    For y = 1 To NumOfRecs
      Get DHandle, CLng(IdxBuff(y)), Emp2Rec
      If DedEmpAmt(x, CLng(IdxBuff(y))) > 0 Then
      TotalDAmt# = OldRound(TotalDAmt# + DedEmpAmt(x, CLng(IdxBuff(y))))
      DedEmpCnt(x) = DedEmpCnt(x) + 1
      SSN = AddDashToSSN(Emp2Rec.EmpSSN)
      RSet DedRpt(1).DAmt = LTrim$(RTrim$(Using(Image1$, DAmt#)))
      RptCnt = RptCnt + 1
      '                                0                                 1                           2                    3                    4                          5
      Print #SubHandle, QPTrim$(ThisDed.DCDESC1); dlm; QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName); dlm; SSN; dlm; DedEmpAmt(x, CLng(IdxBuff(y))); dlm; CStr(x)
      End If
    Next y
    FrmShowPctComp.ShowPctComp x, NumOfDeds
    If FrmShowPctComp.Out = True Then
      Unload frmLoadingRpt
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
    If DedEmpCnt(x) > 0 Then TotActive = TotActive + 1
SkipDed:
  Next x
  Unload FrmShowPctComp
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Close DHandle
  Close THandle
  GoSub SubTotalAll
  '                         0                          1                           2                      3                  4
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(LowDate); dlm; MakeRegDate(HiDate); dlm; TotalDAmt#; dlm; CStr(TotActive)

  Close RHandle
  
  If RptCnt = 0 Then
    MsgBox "No deduction data available for the parameters entered."
    Close
    Unload frmLoadingRpt
    Exit Sub
  End If
  
  arPRDeducAll.Show
  Close
  MainLog ("Payroll Deduction Report processed.")
  
  Exit Sub
  
SubTotalAll:
  Sub2Handle = FreeFile
  Open Sub2RptName$ For Output As Sub2Handle
  For x = 1 To fplistDedNo.SelCount
    Get DedCodeFileHandle, ThisSel(x), ThisDed
      Print #Sub2Handle, QPTrim$(ThisDed.DCDESC1); dlm; Using("#####", DedEmpCnt(ThisSel(x))); dlm; CStr(DedTotAmt(ThisSel(x)))
  Next x
  
  Close Sub2Handle
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPRDeduction", "PrintGraphics", Erl)
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


End Sub

Private Sub LoadThisForm()
   Dim DedCodeHandle As Integer
   Dim DedCodeRec As DedCodeRecType
'   Dim DedCodeRecLen As Integer
   Dim Today As String * 10
   Dim x As Integer
   
   Today = Date '$
   fptxtEndDate.Text = Today
   fptxtStartDate.Text = Mid(Today, 1, 2) + "-" + "01" + "-" + Mid(Today, 7, 4)
   
   OpenDedCodeFile DedCodeHandle
'   DedCodeRecLen = LOF(DedCodeHandle) / Len(DedCodeRec)
   GDedCount = LOF(DedCodeHandle) / Len(DedCodeRec)
'   If DedCodeRecLen = 0 Then
   If GDedCount = 0 Then
     fplistDedNo.AddItem "No records on file."
   Else
'     fplistDedNo.AddItem "ALL"
     For x = 1 To GDedCount 'DedCodeRecLen
       Get DedCodeHandle, x, DedCodeRec
       fplistDedNo.AddItem DedCodeRec.DCDESC1
     Next
   End If
   
   fplistDedNo.ListIndex = 0
   
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"

End Sub


Private Sub fplistDedNo_BeforeDropDown(Cancel As Boolean)
   
End Sub

Private Sub fpList1_Click()

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated through Menu Bar on Payroll Deduction Report.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim FirstTime As Boolean, IdxRecLen As Integer
  Dim FormName As String, x As Integer, y As Integer
  Dim Choice() As String, cnt As Integer, DAmt#
  Dim LowDate As Long, OKFlag As Boolean
  Dim HiDate As Long, IdxFileSize&, TransRecNum&
  Dim EmpRecSize As Long, TempDed$, NumOfRecs As Long
  Dim TRecSize As Long, DNum As Integer, Page As Integer
  Dim DedLineLen As Long, MaxLines As Integer
  Dim DedCodeFile As DedCodeRecType, RecNo As Integer
  Dim Image1$, Image2$, RptTitle$, FF$
  Dim RptName$, RptDate$, RptDed$, UsingThisOne As Boolean
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim EmpData2Name As EmpData2Type, LineCnt As Integer
  Dim TransHistFileName As TransRecType
  Dim TotalDAmt#, UTemp$, UFileHandle As Integer
  Dim DedCodeFileHandle As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim IdxFileName As NameSortIdxType
  'error checking
  Dim Emp2Rec As EmpData2Type
  Dim AllFlag As Boolean
  Dim NumOfDeds As Integer
  Dim GTotalDAmt#, TotEmps As Integer
  Dim EmpCnt As Integer
  Dim ThisDed As DedCodeRecType
  Dim SumFlag As Boolean, LastX As Integer
  Dim EndOfPage As Boolean, TotActive As Integer
  Dim Nextx As Integer
  
  On Error GoTo ERRORSTUFF
  EndOfPage = False
  SumFlag = False
  OpenEmpIdxLNameFile EmpIdxLNameHandle

  NumOfRecs = LOF(EmpIdxLNameHandle) \ 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  Close EmpIdxLNameHandle
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "ERROR: The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  If CheckValDate(fptxtEndDate.Text) = False Then
    MsgBox "ERROR: The End Date is not valid"
    fptxtEndDate.SetFocus
    Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
    MsgBox "ERROR: The Start Date is later than the End Date"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If fplistDedNo.Text = "" Then
    MsgBox "ERROR: Please make a selection in the Deduction Number field"
    fplistDedNo.SetFocus
    Exit Sub
  End If
  If fptxtStartDate.Text = "" Then
    MsgBox "ERROR: Please enter a Starting Date"
    fptxtStartDate.SetFocus
    Exit Sub
  End If

  If fptxtEndDate.Text = "" Then
    MsgBox "ERROR: Please enter an Ending Date"
    fptxtEndDate.SetFocus
    Exit Sub
  End If

  OpenDedCodeFile DedCodeFileHandle
  NumOfDeds = LOF(DedCodeFileHandle) / Len(DedCodeFile)
  ReDim DedCodes(1 To NumOfDeds) As DedCodeRecType
  For cnt = 1 To NumOfDeds ' 50 'load up DedCodes array
    Get DedCodeFileHandle, cnt, DedCodes(cnt)
  Next
  Nextx = 0
  ReDim ThisSel(1 To fplistDedNo.SelCount) As Integer
  For x = 0 To NumOfDeds - 1
    If fplistDedNo.Selected(x) = True Then
      Nextx = Nextx + 1
      ThisSel(Nextx) = x + 1
    End If
  Next x
  
  DNum = ThisSel(Nextx)
  FF$ = Chr$(12)
  Image1$ = "$#,##0.00"
  Image2$ = "$###,##0.00"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim DedRpt(1) As DedRptType

  MaxLines = 55
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  DedLineLen = Len(DedRpt(1))
  OpenUnitFile UFileHandle
  Get UFileHandle, 1, Unit(1)
  IdxRecLen = 2
  
  OpenEmpIdxLNameFile EmpIdxLNameHandle

  NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen

  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  RptTitle$ = "Payroll Deduction Report"

  RptName$ = "PRRPTS\DEDUCT" & DNum & ".RPT"
  RptDate$ = "Start Date: " + MakeRegDate(LowDate) + "  Ending Date: " + MakeRegDate(HiDate) ''Num2Date
  
  If fplistDedNo.SelCount = 1 Then
    RptDed$ = QPTrim$(DedCodes(DNum).DCDESC1) + " Deduction Report." '' + CrLf$
  Else
    RptDed$ = "All Deductions Report"
  End If
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 6, RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  
  x = 1
  GoSub DedHeader
  ReDim DedEmpAmt(1 To NumOfDeds, 1 To NumOfRecs) As Double
  ReDim TotalEmps(1 To NumOfDeds) As Integer
  ReDim DedTotAmt(1 To NumOfDeds) As Double
  ReDim DedEmpCnt(1 To NumOfDeds) As Integer
  frmLoadingRpt.Show
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    'If this employee has not been terminated
    'If there are no transactions for this employee then skip 'em
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm2All
    End If
      TransRecNum& = CLng(Emp2Rec.LastTransRec)
      Do
        Get THandle, TransRecNum&, TransHRec(1)
        Select Case TransHRec(1).CheckDate
        'if this transaction falls within the date parameters
        Case LowDate To HiDate
          'gather all data for this employee and sort into
          'different deductions...Y = Deduction and
          For y = 1 To fplistDedNo.SelCount
            If TransHRec(1).DAmt(ThisSel(y)) <> 0 Then
              DedEmpAmt(ThisSel(y), CLng(IdxBuff(RecNo))) = DedEmpAmt(ThisSel(y), CLng(IdxBuff(RecNo))) + TransHRec(1).DAmt(ThisSel(y))
              DedTotAmt(ThisSel(y)) = OldRound(DedTotAmt(ThisSel(y)) + TransHRec(1).DAmt(ThisSel(y)))
            End If
          Next y
        Case Else
        End Select
        If TransHRec(1).PrevTransRec <= 0 Then
          Exit Do
        Else
          TransRecNum& = CLng(TransHRec(1).PrevTransRec)
        End If
      Loop
SkipEm2All:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      Unload frmLoadingRpt
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  
  Unload FrmShowPctComp
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Nextx = 1
  For x = 1 To NumOfDeds
    If Nextx > fplistDedNo.SelCount Then Exit For
    If ThisSel(Nextx) = x Then
      Nextx = Nextx + 1
    Else
      GoTo NoDed
    End If
    GoSub HeaderAll
    Get DedCodeFileHandle, x, ThisDed
    For y = 1 To NumOfRecs
      Get DHandle, CLng(IdxBuff(y)), Emp2Rec
      Emp2Rec.EmpSSN = AddDashToSSN(Emp2Rec.EmpSSN)
      If DedEmpAmt(x, CLng(IdxBuff(y))) > 0 Then
        TotalEmps(x) = TotalEmps(x) + 1
        TotalDAmt# = OldRound(TotalDAmt# + DedEmpAmt(x, CLng(IdxBuff(y))))
        RSet DedRpt(1).DAmt = LTrim$(RTrim$(Using(Image1$, DAmt#)))
        Print #RHandle, Tab(3); QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName); Tab(37); QPTrim$(Emp2Rec.EmpSSN); Tab(63); Using$("$##,##0.00", DedEmpAmt(x, CLng(IdxBuff(y))))
        LineCnt = LineCnt + 1
        If LineCnt > MaxLines Then
          Print #RHandle, FF$
          EndOfPage = False
          GoSub DedHeader
        End If
      End If
    Next y
    FrmShowPctComp.ShowPctComp x, NumOfDeds
    If FrmShowPctComp.Out = True Then
      Close
      Unload frmLoadingRpt
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
    GoSub TotsAll
    GTotalDAmt# = GTotalDAmt# + TotalDAmt
    TotalDAmt = 0
NoDed:
  Next x
  
  DoEvents
  Unload FrmShowPctComp
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  If fplistDedNo.SelCount > 1 Then
    GoSub DeptSummary
    GoSub GrandTotsAll
  End If

  Close DHandle
  Close THandle
  RPTSetupPRN 123, RHandle '7/24/02
  Close RHandle
  Unload frmLoadingRpt
  ViewPrint RptName$, RptTitle$
  MainLog ("Payroll Deduction Report processed.")

  Exit Sub

DedHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(71)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 62) = "Page:" + Pg(1)
  Print #RHandle, UTemp$
  Print #RHandle, "Payroll Deduction Report"
  Print #RHandle, RptDate$
  If SumFlag = False Then
    Print #RHandle, "   Employee Name                Social Security #       Deduction Amount"
  End If
  Print #RHandle, String$(72, "-")
  LineCnt = 5
  If SumFlag = False Then
    If LastX <> 0 And EndOfPage = False Then 'do not want to reprint
    'a deduction header that would reprint if the call to DedHeader
    'came on the very last maxline
      GoSub HeaderAll
    End If
    LastX = 1
  End If
  
  Return
  
HeaderAll:
  EndOfPage = False
  Print #RHandle, "Deduction Description: "; DedCodes(x).DCDESC1
  Print #RHandle, String$(72, "-")
  Print #RHandle,
  LineCnt = LineCnt + 3
  If LineCnt > MaxLines Then
    Print #RHandle, FF$
    GoSub DedHeader
  End If

  Return

TotsAll:
  RSet DedRpt(1).DAmt = QPTrim$(Using(Image2$, TotalDAmt#))
  If DedRpt(1).DAmt > 0 Then
    TotActive = TotActive + 1
    Print #RHandle,
    Print #RHandle, String$(72, "-")
    Print #RHandle, "Total For "; QPTrim$(DedCodes(x).DCDESC1) + ":"; Tab(29); "# Employees: " + CStr(TotalEmps(x)); Tab(60); Using$("$#,###,##0.00", TotalDAmt#)
    Print #RHandle, String$(72, "-")
    Print #RHandle,
    Print #RHandle,
    LineCnt = LineCnt + 6
  Else
    Print #RHandle, String$(72, "-")
    Print #RHandle, "Total For "; QPTrim$(DedCodes(x).DCDESC1) + ":"; Tab(49); "None"
    Print #RHandle, String$(72, "-")
    Print #RHandle,
    Print #RHandle,
    LineCnt = LineCnt + 5
  End If
  If LineCnt > MaxLines Then
    EndOfPage = True
    Print #RHandle, FF$
    GoSub DedHeader
  End If
Return

DeptSummary:
  SumFlag = True
  x = 1
  Print #RHandle, FF$
  GoSub DedHeader
  Print #RHandle, "Deduction Summary for " + CStr(TotActive) + " Deductions"
  Print #RHandle, String$(72, "-")
  Print #RHandle, Tab(3); "Description"; Tab(33); "# Employees"; Tab(61); "Total Amount "
  Nextx = 1
  For x = 1 To NumOfDeds
    If Nextx > fplistDedNo.SelCount Then Exit For
    If ThisSel(Nextx) = x Then
      Nextx = Nextx + 1
    Else
      GoTo JumpOver
    End If
    Print #RHandle, Tab(3); Tab(3); QPTrim$(DedCodes(x).DCDESC1); Tab(35); Using$("#####", TotalEmps(x)); Tab(60); Using$("$#,###,##0.00", DedTotAmt(x))
    If Nextx = 41 Then
      GoTo NextPage
      Exit For
    End If
JumpOver:
  Next x
  Return
  
NextPage:
  x = 1
  Print #RHandle, FF$
  GoSub DedHeader
  Print #RHandle, "Deduction Summary for " + CStr(TotActive) + " Deductions"
  Print #RHandle, String$(72, "-")
  Print #RHandle, Tab(3); "Description"; Tab(33); "# Employees"; Tab(61); "Total Amount "
  For y = 41 To NumOfDeds
    Print #RHandle, Tab(3); Tab(3); QPTrim$(DedCodes(y).DCDESC1); Tab(35); Using$("#####", TotalEmps(y)); Tab(60); Using$("$#,###,##0.00", DedTotAmt(y))
  Next y
  Return
  
GrandTotsAll:
  LSet DedRpt(1).EmpName = "     Grand Total: "
  LSet DedRpt(1).SSN = ""
  RSet DedRpt(1).DAmt = QPTrim$(Using(Image2$, GTotalDAmt#))
  Print #RHandle, String$(72, "-")
  Print #RHandle,
  Print #RHandle, " Grand Total: "; Tab(59); Using$("$##,###,##0.00", GTotalDAmt#)
  Print #RHandle, FF$
Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPRDeduction", "PrintText", Erl)
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

