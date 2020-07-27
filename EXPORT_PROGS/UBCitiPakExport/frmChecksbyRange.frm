VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmChecksbyNumber 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Checks Issued by Number"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmChecksbyRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5130
      Left            =   2115
      TabIndex        =   3
      Top             =   1867
      Width           =   7410
      _Version        =   196609
      _ExtentX        =   13070
      _ExtentY        =   9049
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
      FrameColor      =   -2147483627
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmChecksbyRange.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3330
         TabIndex        =   2
         Top             =   3330
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
         ColDesigner     =   "frmChecksbyRange.frx":08E6
      End
      Begin VB.CheckBox chkVoidOnly 
         BackColor       =   &H00D0D0D0&
         Caption         =   "List Voided Checks Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   2800
         Width           =   2895
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   370
         Left            =   3600
         TabIndex        =   0
         Top             =   1515
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
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
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   370
         Left            =   3600
         TabIndex        =   1
         Top             =   2190
         Width           =   1830
         _Version        =   196608
         _ExtentX        =   3228
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
         Left            =   4170
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to create the 'Checks Issued by Number' report."
         Top             =   4080
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
         ButtonDesigner  =   "frmChecksbyRange.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1290
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4080
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
         ButtonDesigner  =   "frmChecksbyRange.frx":0DF4
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
         Left            =   1590
         TabIndex        =   7
         Top             =   3435
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Checks Issued by Number"
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
         Left            =   1056
         TabIndex        =   6
         Top             =   624
         Width           =   5388
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
         Left            =   1875
         TabIndex        =   5
         Top             =   1605
         Width           =   1260
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
         Left            =   1920
         TabIndex        =   4
         Top             =   2265
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   828
         Left            =   912
         Top             =   432
         Width           =   5628
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5400
      Left            =   1980
      Top             =   1732
      Width           =   7695
   End
End
Attribute VB_Name = "frmChecksbyNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'8/13 created this new report
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim CancelProcess As Boolean

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmChecksbyNumber
   MainLog ("Checks Issued Report screen exited.")
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
  Call LoadEmpChksIssScreen
  Me.HelpContextID = hlpChecksByNumber
  MainLog ("Checks Issued Report accessed.")
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
  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, ToPrint$
  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$, UTemp$
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UsingThisOne As Boolean
  Dim RecNo As Long, Page As Integer, CrLf$
  Dim Dash As String * 80
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim TransRecLen As Long, TransRecNum&, FF$
  Dim Emp2Rec As EmpData2Type
  Dim TotNumOfChks As Long
  Dim TotAmtOfChks As Double
  Dim TotNegChks As Long
  Dim TotPosChks As Long
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim Image$
  Dim SumTotal As Boolean
  Dim ThisTotal As Double
  Dim ThisCnt As Integer
  Dim TransCnt As Integer
  ReDim ChkIdx(1 To 1) As Long
  ReDim EmpIdx(1 To 1) As String
  ReDim ChkDate(1 To 1) As String
  ReDim Chknum(1 To 1) As String
  ReDim EmpName(1 To 1) As String
  ReDim ChkAmt(1 To 1) As Double
  Dim ChkCnt As Long
  Dim dlm$
  
  dlm$ = "~"
  SumTotal = False
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date"
    fptxtStart.SetFocus
    Exit Sub
  End If

  Image$ = "$###,###,##0.00"
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  EmpRecSize = Len(Emp2Rec)
  
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  frmLoadingRpt.Show

  RptTitle$ = "Employee Checks Issued by Number"
  RptName$ = "PRRPTS\CHKSBYRANGEG.RPT"
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  Call CountandSortChks(TotPosChks, TotNegChks, TotAmtOfChks, TotNumOfChks, ChkCnt, ChkDate(), Chknum(), EmpName(), ChkAmt(), EmpIdx(), ChkIdx(), IdxBuff(), NumOfRecs)
  If CancelProcess = True Then
    Close
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    Unload FrmShowPctComp
    Unload frmLoadingRpt
    Exit Sub
  End If

  GoSub PrintChkInfo
  Close RHandle
  
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  
  Unload frmLoadingRpt
  DoEvents
  
  If ThisCnt = 0 Then
    MsgBox "No checks were found that fit the parameters entered."
    Exit Sub
  End If
  
  arChksByNumRpt.Show
  
  MainLog ("Employee Checks Issued by Number processed.")
  
  Exit Sub

PrintChkInfo:

  For x = 1 To ChkCnt
    If ChkAmt(x) >= 0 And chkVoidOnly.Value = 1 Then GoTo VoidOnly
    ThisCnt = ThisCnt + 1
    '                 0                    1
    Print #RHandle, City; dlm; MakeRegDate(LowDate) + " through " + MakeRegDate(HiDate); dlm;
    '                  2             3             4            5             6                7
    Print #RHandle, ChkDate(x); dlm; Chknum(x); dlm; EmpName(x); dlm; Using("$###,##0.00", ChkAmt(x)); dlm; TotNegChks; dlm; TotPosChks; dlm;
    '                    8                  9                                 6
    Print #RHandle, TotNumOfChks; dlm; Using("$###,###,##0.00", TotAmtOfChks)
VoidOnly:
  Next x
Return


ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload frmLoadingRpt
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub
Private Sub LoadEmpChksIssScreen()
   Dim Today As String * 10
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtStart.Text = "01-01-" + Mid(Today, 7, 4)
   fptxtEnd.Text = Today
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
   CancelProcess = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmChecksbyNumber.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub CountandSortChks(TotPosChks As Long, TotNegChks As Long, TotAmtOfChks As Double, TotNumOfChks As Long, ChkCnt As Long, ChkDate() As String, Chknum() As String, EmpName() As String, ChkAmt() As Double, EmpIdx() As String, ChkIdx() As Long, IdxBuff() As Integer, NumOfRecs As Long)
  Dim Start As Integer
  Dim Last As Integer
  Dim RecNo As Integer
  Dim DHandle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim TransRecNum&
  Dim THandle As Integer
  Dim TransHRec(1) As TransRecType
  Dim x As Integer
  Dim BigNum As Long
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim SmallNum As Long
  Dim ChkCnts$, ChkDates$, ChkAmts$, ChkNums$, EmpNames$
  
  Start = Date2Num(fptxtStart.Text)
  Last = Date2Num(fptxtEnd.Text)
  OpenEmpData2File DHandle
  OpenTransHistFile THandle
  
  FrmShowPctComp.Label1 = "Sorting Checks"

  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  FrmShowPctComp.Show
  DoEvents
 
  'Step 1 - Find out how many checks qualify
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    'this employee has no transactions
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = Emp2Rec.LastTransRec 'get most recent
    'transaction for this employee
    Do
      Get THandle, TransRecNum&, TransHRec(1)
'      If TransRecNum& = 18907 Then Stop
'      TransHRec(1).CheckNum = TransHRec(1).CheckNum
      Select Case TransHRec(1).CheckDate
      'make sure the check for this transaction for
      'falls within the chosen start and end dates
        Case Start To Last 'yep it qualifies
          ChkCnt = ChkCnt + 1 'increment number of qualifying checks
        Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm4:
  FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    CancelProcess = True
    Exit Sub
  End If

  Next RecNo
  If ChkCnt <= 0 Then 'no checks on file
    Unload frmLoadingRpt
    Unload FrmShowPctComp
    MsgBox "No checks on record."
    Close
    CancelProcess = True
    Exit Sub
  End If
  FrmShowPctComp.Label1 = "Sorting Checks"

  FrmShowPctComp.Show
  DoEvents
  
  ReDim ChkIdx(1 To ChkCnt) 'now we can diminsion these arrays
  ReDim EmpIdx(1 To ChkCnt)
  ReDim ChkDate(1 To ChkCnt)
  ReDim Chknum(1 To ChkCnt)
  ReDim EmpName(1 To ChkCnt)
  ReDim ChkAmt(1 To ChkCnt)
  x = 1
'Step 2 - Load up the arrays with the check data
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm5
    End If
    TransRecNum& = Emp2Rec.LastTransRec
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case Start To Last
        If chkVoidOnly.Value = 1 And TransHRec(1).NetPay > 0 Then GoTo VoidSkip
        ChkIdx(x) = TransHRec(1).CheckNum 'This array will be used for sorting
        ChkDate(x) = MakeRegDate(TransHRec(1).CheckDate)
        ChkDates = MakeRegDate(TransHRec(1).CheckDate)
        Chknum(x) = Str(TransHRec(1).CheckNum)
        ChkNums = Str(TransHRec(1).CheckNum)
        EmpName(x) = QPTrim(Emp2Rec.EmpFName) + "    " + QPTrim(Emp2Rec.EmpLName)
        EmpNames = QPTrim(Emp2Rec.EmpFName) + "    " + QPTrim(Emp2Rec.EmpLName)
        ChkAmt(x) = Using("$###,##0.00", TransHRec(1).NetPay)
        ChkAmts = Using("$###,##0.00", TransHRec(1).NetPay)
        'This array holds the data that will appear on the report
        EmpIdx(x) = ChkDates + "  ~  " + ChkNums + " ~  " + EmpNames + " ~ " + ChkAmts
        TotAmtOfChks = TotAmtOfChks + TransHRec(1).NetPay 'this data appears
        'in the report summary as total dollars
        If ChkIdx(x) < 0 Then
          TotNegChks = TotNegChks + 1
        Else
          TotPosChks = TotPosChks + 1
        End If
        TotNumOfChks = TotNumOfChks + 1 'this data will appear
        'in the report summary as the number of checks written
        x = x + 1
      Case Else
      End Select
VoidSkip:
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
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
    CancelProcess = True
    Exit Sub
  End If
  Next RecNo
  Close THandle
  Close DHandle
'Step 3 - shake out the largest check number
  BigNum = 0
  
  For x = 1 To ChkCnt
    If ChkIdx(x) > BigNum Then
      BigNum = ChkIdx(x)
    End If
  Next x
  
  ReDim TempNIdx(1 To ChkCnt) As Long
  ReDim TempEIdx(1 To ChkCnt) As String
  ReDim TempChkDate(1 To ChkCnt) As String
  ReDim TempChkNum(1 To ChkCnt) As String
  ReDim TempEmpName(1 To ChkCnt) As String
  ReDim TempChkAmt(1 To ChkCnt) As Double
 
  Nextx = 1
  SmallNum = BigNum + 1 'To be sure the largest value
  'is included in the sort we now make the largest value
  'one more than the largest value
'Step 4 - Use the largest number to find the smallest
  
  Do
    For x = Nextx To ChkCnt
      If ChkIdx(x) < SmallNum Then 'go thru the array (holding
      'the qualifying checks) and find the smallest number
        SmallNum = ChkIdx(x)
        Thisx = x 'when ChkCnt is reached Thisx is the
        'position holding the smallest check number
      End If
    Next x
    'swap
    TempNIdx(Nextx) = ChkIdx(Nextx) 'take the first check value
    'in the array and assign it to TempNIdx(Nextx)
    TempChkDate(Nextx) = ChkDate(Nextx)
    TempChkNum(Nextx) = Chknum(Nextx)
    TempEmpName(Nextx) = EmpName(Nextx)
    TempChkAmt(Nextx) = ChkAmt(Nextx)
    TempEIdx(Nextx) = EmpIdx(Nextx) 'same for the rest
    'of the data for this check
    ChkIdx(Nextx) = ChkIdx(Thisx) 'now put the smallest value
    'in the next first array position
    ChkDate(Nextx) = ChkDate(Thisx)
    Chknum(Nextx) = Chknum(Thisx)
    EmpName(Nextx) = EmpName(Thisx)
    ChkAmt(Nextx) = ChkAmt(Thisx)
    EmpIdx(Nextx) = EmpIdx(Thisx) 'ditto here
    ChkIdx(Thisx) = TempNIdx(Nextx) 'take the value that
    'was in the first array spot and put it into where
    'the smallest value used to be...if the smallest value was
    'already in the first position then this procedure just
    'swaps itself
    ChkDate(Thisx) = TempChkDate(Nextx)
    Chknum(Thisx) = TempChkNum(Nextx)
    EmpName(Thisx) = TempEmpName(Nextx)
    ChkAmt(Thisx) = TempChkAmt(Nextx)
    EmpIdx(Thisx) = TempEIdx(Nextx) 'ditto
    SmallNum = BigNum + 1 'reset smallnum and start again
    Nextx = Nextx + 1 'the new beginning position is incremented
    'by one and the last position now holds a constant value
    'in sorted order
  Loop Until Nextx > ChkCnt 'stop when the last ChkCnt
  'has been processed
End Sub

Private Sub PrintText()
  Dim LowDate As Long, HiDate As Long
  Dim Image1$, Image2$, ToPrint$, Chknum$
  Dim MaxLines As Integer, LineCnt As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, cnt As Integer
  Dim UnitFileName As Integer, RptTitle$, UTemp$
  Dim RptName$, THandle As Integer, RHandle As Integer
  Dim DHandle As Integer, UsingThisOne As Boolean
  Dim RecNo As Long, Page As Integer, CrLf$
  Dim Dash As String * 80
  Dim City$, UnitHandle As Integer, x As Integer
  Dim UnitFileRec As UnitFileRecType
  Dim EmpIdxLNameHandle As Integer, EmpHistoryRpt$
  Dim TransRecLen As Long, TransRecNum&, FF$
  Dim Emp2Rec As EmpData2Type
  Dim TotNumOfChks As Integer
  Dim TotAmtOfChks As Double
  Dim TotNegChks As Double
  Dim TotPosChks As Double
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim Image$
  Dim SumTotal As Boolean
  Dim ThisTotal As Double
  Dim ThisCnt As Integer
  Dim TransCnt As Integer
  ReDim ChkIdx(1 To 1) As Long
  ReDim EmpIdx(1 To 1) As String
  Dim ChkCnt As Long
  
  SumTotal = False
  FF$ = Chr(12)
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  If HiDate < LowDate Then
    MsgBox "ERROR: The End Date is earlier than the Start Date"
    fptxtStart.SetFocus
    Exit Sub
  End If

  Image$ = "$###,###,##0.00"
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle

  ToPrint$ = Space$(78)
  Chknum$ = Space$(12)

  EmpRecSize = Len(Emp2Rec)
  
  MaxLines = 57
  LineCnt = 0
  Dash = String$(78, "-") + CrLf$

  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  OpenEmpIdxLNameFile EmpIdxLNameHandle
  
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  Call CountandSortChksText(TotPosChks, TotNegChks, TotAmtOfChks, TotNumOfChks, ChkCnt, EmpIdx(), ChkIdx(), IdxBuff(), NumOfRecs)
  
  RptTitle$ = "Employee Checks Issued by Number"
  RptName$ = "PRRPTS\CHKSBYRANGE.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  GoSub PrintIssuedHeader
  GoSub PrintChkInfo
  Print #RHandle, FF$
  GoSub PrintChkTotals
  Print #RHandle, FF$
  Close RHandle
  Unload frmLoadingRpt
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  
  If ThisCnt = 0 Then
    MsgBox "No checks were found that fit the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  MainLog ("Employee Checks Issued by Number processed.")
  
  Exit Sub

PrintIssuedHeader:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = City
  Mid$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  Print #RHandle, UTemp$
  Print #RHandle, "Employee Checks Issued by Number" + CrLf$
  Print #RHandle, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate) + CrLf$
  If SumTotal = False Then
    Print #RHandle, "Check Date    Check Number        Employee Name                       Net Pay" + CrLf$
    Print #RHandle, Dash
  End If
  LineCnt = 5
  Return

PrintChkInfo:
  For x = 1 To ChkCnt
    ThisCnt = ThisCnt + 1
    Print #RHandle, EmpIdx(x)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines Then            'bottom of page?
      Print #RHandle, FF$   'yes; form feed
      GoSub PrintIssuedHeader
    End If
    FrmShowPctComp.ShowPctComp x, ChkCnt
  Next x
Return

PrintChkTotals:
  Page = Page + 1
  RSet Pg(1) = Page
  UTemp$ = Space$(80)
  LSet UTemp$ = City
  Mid$(UTemp$, 71) = "Page:" + Pg(1) + CrLf$
  Print #RHandle, "Employee Checks Issued by Number" + CrLf$
  Print #RHandle, MakeRegDate(LowDate) + " thru " + MakeRegDate(HiDate) + CrLf$
  Print #RHandle, UTemp$
  Print #RHandle,
  Print #RHandle, Dash
  Print #RHandle, Tab(8); "Total Negative Checks: "; Tab(35); Using("#######", Str(TotNegChks))
  Print #RHandle, Tab(8); "Total Positive Checks: "; Tab(35); Using("#######", Str(TotPosChks))
  Print #RHandle,
  Print #RHandle, Tab(8); "Total Check Entries: "; Tab(35); Using("#######", Str(TotNumOfChks))
  Print #RHandle,
  Print #RHandle, Tab(8); "Total Amount of Checks: "; Tab(35); Using(Image, TotAmtOfChks)
  

Return

End Sub

Private Sub CountandSortChksText(TotPosChks As Double, TotNegChks As Double, TotAmtOfChks As Double, TotNumOfChks As Integer, ChkCnt As Long, EmpIdx() As String, ChkIdx() As Long, IdxBuff() As Integer, NumOfRecs As Long)
  Dim Start As Integer
  Dim Last As Integer
  Dim RecNo As Integer
  Dim DHandle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim TransRecNum&
  Dim THandle As Integer
  Dim TransHRec(1) As TransRecType
  Dim x As Integer
  Dim BigNum As Long
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim SmallNum As Long
  Dim ChkDate As String * 12
  Dim Chknum As String * 12
  Dim EmpName As String * 32
  Dim ChkAmt As String * 12
  
  Start = Date2Num(fptxtStart.Text)
  Last = Date2Num(fptxtEnd.Text)
  OpenEmpData2File DHandle
  OpenTransHistFile THandle
  
  FrmShowPctComp.Label1 = "Employee Checks by Number"
  
  frmLoadingRpt.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  'Step 1 - Find out how many checks qualify
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    'this employee has no transactions
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm4
    End If
    TransRecNum& = Emp2Rec.LastTransRec 'get most recent
    'transaction for this employee
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      If chkVoidOnly.Value = 1 And TransHRec(1).NetPay > 0 Then GoTo VoidOnly
      Select Case TransHRec(1).CheckDate
      'make sure the check for this transaction for
      'falls within the chosen start and end dates
      Case Start To Last 'yep it qualifies
        ChkCnt = ChkCnt + 1 'increment number of qualifying checks
      Case Else
      End Select
VoidOnly:
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm4:
  Next RecNo
  If ChkCnt <= 0 Then 'no checks on file
    Unload frmLoadingRpt
    Unload FrmShowPctComp
    MsgBox "No checks on record."
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Show , Me
  DoEvents
  
  ReDim ChkIdx(1 To ChkCnt) 'now we can diminsion these arrays
  ReDim EmpIdx(1 To ChkCnt)
  x = 1
'Step 2 - Load up the arrays with the check data
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm5
    End If
    TransRecNum& = Emp2Rec.LastTransRec
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case Start To Last
        If chkVoidOnly.Value = 1 And TransHRec(1).NetPay > 0 Then GoTo SkipVoid
        ChkIdx(x) = TransHRec(1).CheckNum 'This array will be used for sorting
        LSet ChkDate = MakeRegDate(TransHRec(1).CheckDate)
        LSet Chknum = Str(TransHRec(1).CheckNum)
        LSet EmpName = QPTrim(Emp2Rec.EmpFName) + " " + QPTrim(Emp2Rec.EmpLName)
        RSet ChkAmt = Using("$###,##0.00", TransHRec(1).NetPay)
        'This array holds the data that will appear on the report
        EmpIdx(x) = ChkDate + "    " + Chknum + "   " + EmpName + "  " + ChkAmt
        TotAmtOfChks = TotAmtOfChks + TransHRec(1).NetPay 'this data appears
        'in the report summary as total dollars
        If ChkIdx(x) < 0 Then
          TotNegChks = TotNegChks + 1
        Else
          TotPosChks = TotPosChks + 1
        End If
        TotNumOfChks = TotNumOfChks + 1 'this data will appear
        'in the report summary as the number of checks written
        x = x + 1
      Case Else
      End Select
SkipVoid:
      If TransHRec(1).PrevTransRec <= 0 Then    'if no more transactions
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEm5:
  FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
  Next RecNo
  Close THandle
  Close DHandle
'Step 3 - shake out the largest check number
  BigNum = 0
  For x = 1 To ChkCnt
    If ChkIdx(x) > BigNum Then
      BigNum = ChkIdx(x)
    End If
  FrmShowPctComp.ShowPctComp x, ChkCnt
  Next x
  
  ReDim TempNIdx(1 To ChkCnt) As Long
  ReDim TempEIdx(1 To ChkCnt) As String
 
  Nextx = 1
  SmallNum = BigNum + 1 'To be sure the largest value
  'is included in the sort we now make the largest value
  'one more than the largest value
'Step 4 - Use the largest number to find the smallest
  Do
    For x = Nextx To ChkCnt
      If ChkIdx(x) < SmallNum Then 'go thru the array (holding
      'the qualifying checks) and find the smallest number
        SmallNum = ChkIdx(x)
        Thisx = x 'when ChkCnt is reached Thisx is the
        'position holding the smallest check number
      End If
    Next x
    'swap
    TempNIdx(Nextx) = ChkIdx(Nextx) 'take the first check value
    'in the array and assign it to TempNIdx(Nextx)
    TempEIdx(Nextx) = EmpIdx(Nextx) 'same for the rest
    'of the data for this check
    ChkIdx(Nextx) = ChkIdx(Thisx) 'now put the smallest value
    'in the next first array position
    EmpIdx(Nextx) = EmpIdx(Thisx) 'ditto here
    ChkIdx(Thisx) = TempNIdx(Nextx) 'take the value that
    'was in the first array spot and put it into where
    'the smallest value used to be...if the smallest value was
    'already in the first position then this procedure just
    'swaps itself
    EmpIdx(Thisx) = TempEIdx(Nextx) 'ditto
    SmallNum = BigNum + 1 'reset smallnum and start again
    FrmShowPctComp.ShowPctComp Nextx, ChkCnt
    Nextx = Nextx + 1 'the new beginning position is incremented
    'by one and the last position now holds a constant value
    'in sorted order
  Loop Until Nextx > ChkCnt 'stop when the last ChkCnt
  'has been processed
  
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

