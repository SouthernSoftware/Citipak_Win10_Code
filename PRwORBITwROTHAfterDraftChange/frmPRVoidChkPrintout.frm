VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPRVoidChkPrintout 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll: Void Check Data Review"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   Icon            =   "frmPRVoidChkPrintout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8595
   StartUpPosition =   1  'CenterOwner
   Begin LpLib.fpList fpList 
      Height          =   2910
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   7575
      _Version        =   196608
      _ExtentX        =   13361
      _ExtentY        =   5133
      TextAlias       =   ""
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmPRVoidChkPrintout.frx":08CA
   End
   Begin EditLib.fpText fptxtHeader 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   7575
      _Version        =   196608
      _ExtentX        =   13361
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      AutoAdvance     =   0   'False
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
   Begin EditLib.fpText fptxtChoice 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
      _Version        =   196608
      _ExtentX        =   661
      _ExtentY        =   238
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      Text            =   "fpText1"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAbort 
      Height          =   570
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2595
      _Version        =   131072
      _ExtentX        =   4577
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
      ButtonDesigner  =   "frmPRVoidChkPrintout.frx":0CCA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdVoid 
      Height          =   570
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2595
      _Version        =   131072
      _ExtentX        =   4577
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
      ButtonDesigner  =   "frmPRVoidChkPrintout.frx":0EE4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   570
      Left            =   3480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1515
      _Version        =   131072
      _ExtentX        =   2672
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
      ButtonDesigner  =   "frmPRVoidChkPrintout.frx":10C5
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPRVoidChkPrintout.frx":12D9
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5775
      Left            =   240
      Top             =   240
      Width           =   8070
   End
End
Attribute VB_Name = "frmPRVoidChkPrintout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim TransDesc() As String
  Dim TransAmt() As String
  Dim TransType() As String
  Dim TransCnt As Integer
  Dim ThisEmpName As String
  Dim ThisTown As String
  Dim ThisCheckNumber As String

Private Sub cmdVoid_Click()
  Me.Hide
  fptxtChoice = "continue"
End Sub

Private Sub cmdAbort_Click()
  Me.Hide
  fptxtChoice = "abort"
End Sub

Private Sub cmdPrint_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PrintText
  ElseIf RptOpt = 1 Then
    Call PrintGraphics
  Else
    Exit Sub
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyEscape:
    Call cmdAbort_Click
    KeyCode = 0
  Case vbKeyF10:
    Call cmdVoid_Click
    KeyCode = 0
  Case vbKeyF3:
    Call cmdPrint_Click
    KeyCode = 0
  Case Else:
  End Select
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim RptHandle As Integer
  Dim x As Integer
  
  ReportFile$ = "VoidPrnt.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  
  Print #RptHandle, "TRANSACTION ENTRIES"
  Print #RptHandle, String$(62, "-")
  LineCnt = LineCnt + 2
  Dim PrintSecondGroup As Boolean
  
  PrintSecondGroup = False
  For x = 2 To TransCnt
    If TransType(x) <> "TRANSACTION ENTRIES" And PrintSecondGroup = False Then
      If LineCnt > MaxLines - 5 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      Print #RptHandle, ""
      Print #RptHandle, ""
      Print #RptHandle, TransType(x)
      Print #RptHandle, String$(62, "-")
      LineCnt = LineCnt + 4
      PrintSecondGroup = True
    End If
    Print #RptHandle, TransDesc(x); Tab(50); TransAmt(x)
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  Next x
  
  Print #RptHandle, FF$
  
  Close RptHandle
  
  ViewPrint ReportFile$, "Void Check Data", True, , False
  
  KillFile ReportFile$
  
  If PrintSecondGroup = True Then
    MainLog ("User printed text version of 'Void Check Data Review' for check # " + ThisCheckNumber + " with transaction entries and GL entries printed.")
  Else
    MainLog ("User printed text version of 'Void Check Data Review' for check # " + ThisCheckNumber + " with transaction entries only. No GL entries available.")
  End If
  
  Exit Sub

PrintHeader:
  Print #RptHandle, Tab(15); "Void Payroll Check Data Review"
  Print #RptHandle, ThisTown
  Print #RptHandle, ThisEmpName; Tab(40);
  Print #RptHandle, "Check Number: " + ThisCheckNumber
  Print #RptHandle,
  Print #RptHandle, String$(62, "=")
  LineCnt = 5

  Return

End Sub

Private Sub PrintGraphics()
  Dim x As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim dlm$
  
  dlm$ = "~"
  ReportFile$ = "PRRPTS\VOIDPRN.RPT"  'Report File Name
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  For x = 2 To TransCnt
    Print #RptHandle, ThisEmpName; dlm; ThisTown; dlm; ThisCheckNumber; dlm;
    Print #RptHandle, TransDesc(x); dlm; TransAmt(x); dlm; TransType(x)
  Next x
  
  Close RptHandle
  
  arPRVoidChkPrint.Show vbModal
  
  
End Sub

Private Sub LoadMe()
  Dim cnt As Integer
  Dim TransRec As TransRecType
  Dim CheckNum As Long
  Dim CheckDate As Integer
  Dim x As Double
  Dim VoidRec As VoidCheckType
  Dim VHandle As Integer
  Dim VoidCnt As Double
  Dim THandle As Integer
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim NumOfTempVoids As Integer
  Dim VPHandle As Integer 'Void Post
  Dim NumOfVoidPosts As Double
  Dim VCnt As Integer
  Dim Acct As GLAcctRecType
  Dim AcctRecLen As Long
  Dim NumAccts As Long, NumTrans&
  Dim AcctFileNum As Integer
  Dim TransFileNum As Integer
  Dim Tran2Post As GLTransRecType
  Dim FundLen As Integer
  Dim DeptLen As Integer
  Dim DetLen As Integer, DedCnt As Integer
  Dim SysRec As RegDSysFileRecType
  Dim SHandle As Integer, RecdNum#
  Dim Source$, AcctIndexName$
  Dim EmpRec2 As EmpData2Type
  Dim EHandle As Integer
  Dim ThisEmpNum$
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim ThisTransCnt As Integer
  
  TransCnt = 0
  OpenUnitFile UHandle
  Get UHandle, 1, UnitRec
  Close UHandle
  ThisTown = QPTrim$(UnitRec.UFCITY)
  
  OpenSysFile SHandle
  Get SHandle, 1, SysRec
  Close SHandle
  
'  Call GetAcctStruct(QPTrim$(SysRec.CITIDIR), FundLen, DeptLen, DetLen)
  Call GetAcctStruct(CurrCitiPath, FundLen, DeptLen, DetLen)
  
  If GlobalTransNum = 0 Then Exit Sub 'may want to add a message here
  
  OpenTransHistFile THandle
  Get THandle, GlobalTransNum, TransRec
  Close THandle
  
  OpenEmpData2File EHandle
  Get EHandle, TransRec.EmpPin, EmpRec2
  Close EHandle
  ThisEmpNum$ = QPTrim$(EmpRec2.EmpNo)
  ThisEmpName$ = QPTrim$(EmpRec2.EmpFName) + " " + QPTrim$(EmpRec2.EmpLName)
  
  fptxtHeader.Text = GlobalName
  CheckNum = TransRec.CheckNum
  ThisCheckNumber = TransRec.CheckNum
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Check Number"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.CheckNum)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Check Date"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = MakeRegDate(TransRec.CheckDate)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  CheckDate = TransRec.CheckDate
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Regular Hours Worked"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.RegHrsWork)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Overtime Hours Worked"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.OTHours)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Vacation Hours Used"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.VacUsed)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Sick Hours Used"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.SickUsed)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Holiday Hours Used"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.HOLHOURS)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Personal Hours Used"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.PerHours)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Comp Hours Used"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.CompUsed)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Regular Hours Paid"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.RegHrsPaid)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Overtime Hours Paid"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.OTHrsPaid)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Overtime to Comp"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("#########0", TransRec.OT2Comp)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  For cnt = 1 To 8
    If TransRec.TDist(cnt).DRHrs > 0 Then
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = TransRec.TDist(cnt).DAcct + "  Reg Hours"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("#########0", TransRec.TDist(cnt).DRHrs)
      fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
      
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = TransRec.TDist(cnt).DAcct + "  Wage"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TDist(cnt).DRWage)
      fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & Using$("$##,##0.00", TransAmt(TransCnt))
    End If
    
    If TransRec.TDist(cnt).DOHrs > 0 Then
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = TransRec.TDist(cnt).DAcct + "  OT Hours"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("#########0", TransRec.TDist(cnt).DOHrs)
      fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
      
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = TransRec.TDist(cnt).DAcct + "  Wage"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TDist(cnt).DOWage)
      fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & Using$("$##,##0.00", TransAmt(TransCnt))
    End If
  Next
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Total Reg Wage"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TotRegWage)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Total OT Wage"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TotOTWage)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Gross Wage"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.GrossWage)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  For cnt = 1 To 3
    If TransRec.EAmt(cnt) > 0 Then
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = "Alternate Earnings"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.EAmt(cnt))
      fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
    End If
  Next
  
  For cnt = 1 To 6
    If TransRec.EDist(cnt).EAmt > 0 Then
      TransCnt = TransCnt + 1
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = QPTrim$(TransRec.EDist(cnt).EAcct) + " Alt Earn"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.EDist(cnt).EAmt)
      fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
    End If
  Next
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Total Additional Earnings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TotAdditEarn)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.GrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Social Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.SocGrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Medicare Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.MedGrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Federal Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.FedGrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "State Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.StaGrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Social Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.SocTaxAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Medicare Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.MedTaxAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Federal Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.FedTaxAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "State Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.StaTaxAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Retirement Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.RetireAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Total Withholdings"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TotTaxAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  For cnt = 1 To 50
    If TransRec.DAmt(cnt) > 0 Then
      ReDim Preserve TransDesc(1 To TransCnt) As String
      TransDesc(TransCnt) = "Deduction"
      ReDim Preserve TransAmt(1 To TransCnt) As String
      TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.DAmt(cnt))
      fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
    End If
  Next
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Total Deductions"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TotDedAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "EIC Amount"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.EICAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Net Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.NetPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Matching Retirement Amount"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.MatchRetAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Matching Social Amount"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.MatchSocAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Matching Medicare Amount"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.MatchMedAmt)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Retirement Gross Pay"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.RetGrossPay)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  TransCnt = TransCnt + 1
  ReDim Preserve TransDesc(1 To TransCnt) As String
  TransDesc(TransCnt) = "Taxable Fringe"
  ReDim Preserve TransAmt(1 To TransCnt) As String
  TransAmt(TransCnt) = Using$("$##,##0.00", TransRec.TaxFring)
  fpList.InsertRow = TransDesc(TransCnt) & Chr(9) & TransAmt(TransCnt)
  
  ThisTransCnt = TransCnt
  ReDim TransType(1 To TransCnt) As String
  For x = 1 To TransCnt
    TransType(x) = "TRANSACTION ENTRIES"
  Next x

  OpenVoidChkPostFile VHandle
  VoidCnt = LOF(VHandle) / Len(VoidRec)

  OpenGLAcctFile AcctFileNum 'GLACCT.DAT
  NumAccts = LOF(AcctFileNum) / Len(Acct)

  OpenGLTransFile TransFileNum 'GLTRANS.DAT
  NumTrans& = LOF(TransFileNum) / Len(Tran2Post)

  For x = 1 To VoidCnt
    Get VHandle, x, VoidRec
      If VoidRec.CheckNum = CheckNum Then
        If VoidRec.CheckDate = CheckDate Then
          If QPTrim$(VoidRec.EmpNum) = ThisEmpNum$ Then
            If VoidRec.FEDWHAmt = 0 Then GoTo Skip1
            VoidRec.FEDWHGL = AddDashesToGLNumber(VoidRec.FEDWHGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.FEDWHGL)  'Verify account is in G/L
            If RecdNum > 0 Then                  'if valid acct then proceed
              Get AcctFileNum, RecdNum, Acct    'Get the account
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.FEDWHGL) + "  VPFEDWH"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.FEDWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.FEDWHGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.FEDWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip1:
            If VoidRec.MEDMATCRAmt = 0 Then GoTo Skip2
            VoidRec.MEDMATCRGL = AddDashesToGLNumber(VoidRec.MEDMATCRGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDMATCRGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDMATCRGL) + "  VPMEDMATCHCRE"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDMATCRGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip2:
            If VoidRec.MEDMATDBAmt = 0 Then GoTo Skip3
            VoidRec.MEDMATDBGL = AddDashesToGLNumber(VoidRec.MEDMATDBGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDMATDBGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDMATDBGL) + "  VPMEDMATCHDBT"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDMATDBGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip3:
            If VoidRec.MEDWHAmt = 0 Then GoTo Skip4
            VoidRec.MEDWHGL = AddDashesToGLNumber(VoidRec.MEDWHGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDWHGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDWHGL) + "  VPMEDWH"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.MEDWHGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.MEDWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip4:
            If VoidRec.PRNet = 0 Then GoTo Skip5
            VoidRec.PRNetGL = AddDashesToGLNumber(VoidRec.PRNetGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.PRNetGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PRNetGL) + "  VPPRNET"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PRNet)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PRNetGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PRNet)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip5:
            If VoidRec.RETMATCRAmt = 0 Then GoTo Skip6
            VoidRec.RETMATCRGL = AddDashesToGLNumber(VoidRec.RETMATCRGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETMATCRGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETMATCRGL) + "  VPRETMATCHCRE"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETMATCRGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip6:
            If VoidRec.RETMATDBAmt = 0 Then GoTo Skip7
            VoidRec.RETMATDBGL = AddDashesToGLNumber(VoidRec.RETMATDBGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETMATDBGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETMATDBGL) + "  VPRETMATCHDBT"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETMATDBGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip7:
            If VoidRec.RETWHAmt = 0 Then GoTo Skip8
            VoidRec.RETWHGL = AddDashesToGLNumber(VoidRec.RETWHGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETWHGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETWHGL) + "  VPRETWH"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.RETWHGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.RETWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip8:
            If VoidRec.SOCMATCRAmt = 0 Then GoTo Skip9
            VoidRec.SOCMATCRGL = AddDashesToGLNumber(VoidRec.SOCMATCRGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCMATCRGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCMATCRGL) + "  VPSOCMATCHCRE"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCMATCRGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCMATCRAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip9:
            If VoidRec.SOCMATDBAmt = 0 Then GoTo Skip10
            VoidRec.SOCMATDBGL = AddDashesToGLNumber(VoidRec.SOCMATDBGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCMATDBGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCMATDBGL) + "  VPSOCMATCHDBT"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCMATDBGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCMATDBAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip10:
            If VoidRec.SOCWHAmt = 0 Then GoTo Skip11
            VoidRec.SOCWHGL = AddDashesToGLNumber(VoidRec.SOCWHGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCWHGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCWHGL) + "  VPSOCWH"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.SOCWHGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.SOCWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip11:
            If VoidRec.STAWHAmt = 0 Then GoTo Skip12
            VoidRec.STAWHGL = AddDashesToGLNumber(VoidRec.STAWHGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.STAWHGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.STAWHGL) + "  VPSTAWH"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.STAWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.STAWHGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.STAWHAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip12:
            If VoidRec.WagesAmt = 0 Then GoTo Skip13
            VoidRec.WagesGL = AddDashesToGLNumber(VoidRec.WagesGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.WagesGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.WagesGL) + "  VPWAGES"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.WagesAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.WagesGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.WagesAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip13:
            If VoidRec.PPEAmt = 0 Then GoTo Skip14
            VoidRec.PPEGL = AddDashesToGLNumber(VoidRec.PPEGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.PPEGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PPEGL) + "  VPCENTPPE"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PPEAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PPEGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PPEAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip14:
            If VoidRec.PPETotAmt = 0 Then GoTo Skip15
            VoidRec.PPETotGL = AddDashesToGLNumber(VoidRec.PPETotGL, FundLen, DeptLen, DetLen)
            RecdNum = FindAcct(AcctIndexName$, VoidRec.PPETotGL)
            If RecdNum > 0 Then
              Get AcctFileNum, RecdNum, Acct
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PPETotGL) + "  VPCENTPPETOTAL"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PPETotAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
            Else
              TransCnt = TransCnt + 1
              ReDim Preserve TransDesc(1 To TransCnt) As String
              TransDesc(TransCnt) = QPTrim$(VoidRec.PPETotGL) + "  ***ERROR***"
              ReDim Preserve TransAmt(1 To TransCnt) As String
              TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.PPETotAmt)
              fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
            End If

Skip15:
            For DedCnt = 1 To 50
              If VoidRec.DedData(DedCnt).DAmt <> 0 Then
                VoidRec.DedData(DedCnt).DedGLNum = AddDashesToGLNumber(VoidRec.DedData(DedCnt).DedGLNum, FundLen, DeptLen, DetLen)
                RecdNum = FindAcct(AcctIndexName$, VoidRec.DedData(DedCnt).DedGLNum)
                If RecdNum > 0 Then
                  Get AcctFileNum, RecdNum, Acct
                  TransCnt = TransCnt + 1
                  ReDim Preserve TransDesc(1 To TransCnt) As String
                  TransDesc(TransCnt) = QPTrim$(VoidRec.DedData(DedCnt).DedGLNum) + "  " + QPTrim$(VoidRec.DedData(DedCnt).DedDesc)
                  ReDim Preserve TransAmt(1 To TransCnt) As String
                  TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.DedData(DedCnt).DAmt)
                  fpList.InsertRow = TransDesc(TransCnt) + Chr(9) & TransAmt(TransCnt)
                Else
                  TransCnt = TransCnt + 1
                  ReDim Preserve TransDesc(1 To TransCnt) As String
                  TransDesc(TransCnt) = QPTrim$(VoidRec.DedData(DedCnt).DedGLNum) + "  ***ERROR***"
                  ReDim Preserve TransAmt(1 To TransCnt) As String
                  TransAmt(TransCnt) = Using$("$##,##0.00", VoidRec.DedData(DedCnt).DAmt)
                  fpList.InsertRow = TransDesc(TransCnt) + Chr(9) + TransAmt(TransCnt)
                End If
              End If
            Next DedCnt
          End If
        End If
      End If
  Next x
  
  If TransCnt = ThisTransCnt Then
    TransCnt = TransCnt + 1
    ReDim Preserve TransType(1 To TransCnt) As String
    ReDim Preserve TransAmt(1 To TransCnt) As String
    ReDim Preserve TransDesc(1 To TransCnt) As String
    TransAmt(TransCnt) = ""
    TransDesc(TransCnt) = ""
    TransType(TransCnt) = "NO GL ENTRIES"
  ElseIf TransCnt > ThisTransCnt Then
    ReDim Preserve TransType(1 To TransCnt) As String
    ThisTransCnt = ThisTransCnt + 1
    For x = ThisTransCnt To TransCnt
      TransType(x) = "GL ENTRIES"
    Next x
  End If
    
    
    
'  For X = 1 To TransCnt
'    Debug.Print TransDesc(X) + "  " + TransAmt(X)
'  Next X
End Sub

Private Sub Form_Load()
  Call LoadMe
End Sub
