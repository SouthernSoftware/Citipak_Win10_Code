VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmORBITPost 
   BackColor       =   &H008F8265&
   Caption         =   "Create Export File"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   Icon            =   "frmORBITPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   495
      Left            =   4418
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   3435
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
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
      ButtonDesigner  =   "frmORBITPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCreate 
      Height          =   495
      Left            =   4418
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   4155
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
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
      ButtonDesigner  =   "frmORBITPost.frx":0AAD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4418
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   4875
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
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
      ButtonDesigner  =   "frmORBITPost.frx":0C97
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      Height          =   3375
      Left            =   2978
      Top             =   2715
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   855
      Index           =   1
      Left            =   1485
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ORBIT: Create Export Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2925
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1485
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frmORBITPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim NameIdx() As Integer
Dim NameCnt As Integer

Private Sub cmdCreate_Click()
  Dim OERec As OrbitEmpData
  Dim OEHandle As Integer
  Dim NumOfOERecs As Integer
  Dim OHRec As OrbitHeader
  Dim OHHandle As Integer
  Dim ODRec As OrbitDetail
  Dim ODHandle As Integer
  Dim NumOfODRecs As Integer
  Dim OTRec As OrbitTrailer
  Dim OTHandle As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim RptFileName As String
  Dim x As Integer, y As Integer, t As Integer
  Dim TSalary As Double
  Dim TempCont As Double
  Dim TCnt As Integer
  Dim DeptNum As String
  Dim TotOReg As Double
  Dim TotReg As Double
  Dim TotOOT As Double
  Dim TotOT As Double
  Dim TotOEmployeeMatch As Double
  Dim TotOEmployerMatch As Double
  Dim TotEmployeeMatch As Double
  Dim TotEmployerMatch As Double
  Dim TransRec As TransRecType
  Dim THandle As Integer
  Dim NextRec As Long
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim ODCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  frmMessageWOpts.Label1.Caption = "Creating the NC ORBIT File will automatically delete the temporary employee pay file for the pay date currently entered. This file can be recreated by re-processing and re-editing as necessary. Do you wish to continue?"
  frmMessageWOpts.Label1.Top = 650
  frmMessageWOpts.cmdCont.Text = "F10 Continue"
  frmMessageWOpts.cmdExit.Text = "ESC Abort"
  frmMessageWOpts.Show vbModal
  If frmMessageWOpts.fptxtChoice.Text = "abort" Then
    Unload frmMessageWOpts
    Exit Sub
  End If
  
  If Not Exist("orbpostprint.dat") Then
    frmMessageWOpts.Label1.Caption = "It is highly recommended that you print a copy of the latest ORBIT transactions before creating the export files. Would you like to print a report now?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Print Report"
    frmMessageWOpts.cmdExit.Text = "ESC Don't Print"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      MainLog ("Recommendation to print the ORBIT report before creating export file was not exercised.")
    Else
      Unload frmMessageWOpts
      frmReportOpt.Show vbModal
      If RptOpt = 2 Then
        MsgBox ("The printer pitch should be no less than pitch 12.")
        Call PrintText
        MainLog ("User printed ORBIT report in text.")
        Exit Sub
      ElseIf RptOpt = 1 Then
        Call PrintGraphics
        MainLog ("User printed ORBIT report in graphics.")
      Else
        Exit Sub
      End If
      MainLog ("User printed ORBIT report right before creating export file.")
    End If
  Else
   KillFile "orbpostprint.dat"
  End If
  
  OpenOrbHeader OHHandle
  Get OHHandle, 1, OHRec
  Close OHHandle
  OpenOrbDetail ODHandle, NumOfODRecs

  GoSub CheckTotals
 
  RptFileName = OHRec.RptPeriod
  ReportFile$ = "NCORBIT/" & RptFileName & OHRec.AgencyNum & ".CNT"
  
  If Exist(ReportFile$) Then
    KillFile ReportFile$
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
'  y = 0
'  t = 0
'  For x = 1 To 540
'    y = y + 1
'    If y = 100 Then
'    y = 1
'      t = t + 1
''      Print #RptHandle, Mid(CStr(x + 1), 1, 1);
'      Print #RptHandle, CStr(t);
'    End If
'    If x = 540 Then
'      Print #RptHandle, " "
'    Else
'      Print #RptHandle, " ";
'    End If
'  Next x
'
'  For x = 1 To 54
'    For y = 1 To 10
'      t = y
'      If y = 10 Then t = 0
'      If y = 10 And x = 54 Then
'        Print #RptHandle, CStr(t)
'      Else
'        Print #RptHandle, CStr(t);
'      End If
'    Next y
'  Next x
  
  Print #RptHandle, OHRec.RecType;
  Print #RptHandle, RightJustNumeric(OHRec.AgencyNum, 5, -1);
  Print #RptHandle, OHRec.RptPeriod;
  Print #RptHandle, OHRec.FrmtVersion;
  Print #RptHandle, FormatThisPayPd(MakeRegDate(OHRec.FileCreateDate), 2)
  
  For x = 1 To NumOfODRecs
    Get ODHandle, x, ODRec
    ODRec.PostedYN = "Y"
    Put ODHandle, x, ODRec
    If ODRec.Deleted = True Then GoTo Deleted
    If QPTrim$(ODRec.TermType) = "NA" Then
      ODRec.TermType = ""
    End If
    If QPTrim$(ODRec.Adjustment) = "NA" Then
      ODRec.Adjustment = ""
    End If
    If QPTrim$(ODRec.TerminationDate) = "0" Then
      ODRec.TerminationDate = ""
    Else
      OpenOrbEmpData OEHandle, NumOfOERecs
      For y = 1 To NumOfOERecs
        Get OEHandle, y, OERec
        If QPTrim$(OERec.EmpNum) = QPTrim$(ODRec.EmpNum) Then
          OERec.Deleted = True
          Put OEHandle, y, OERec
        End If
      Next y
      Close OEHandle
    End If
    If QPTrim$(ODRec.ContrPdEmpBegDate) = "0" Then
      ODRec.ContrPdEmpBegDate = ""
    End If
    If QPTrim$(ODRec.ContrPdEmpEndDate) = "0" Then
      ODRec.ContrPdEmpEndDate = ""
    End If
    If QPTrim$(ODRec.DeptNum) <> "" Then
      DeptNum = RightJustNumeric(ODRec.DeptNum, 6, -1)
    Else
      DeptNum = ODRec.DeptNum
    End If
    
    ODRec.RecType = "D"
    
    Print #RptHandle, ODRec.RecType; RightJustNumeric(OHRec.AgencyNum, 5, -1); LeftJustAlphaNumeric(ODRec.PlanCode, 10); RightJustNumeric(ODRec.SSN, 9, -1);
    Print #RptHandle, RightJustNumeric(ODRec.MemberID, 9, -1); LeftJustAlphaNumeric(ODRec.FirstName, 50); LeftJustAlphaNumeric(ODRec.MiddleName, 50);
    Print #RptHandle, LeftJustAlphaNumeric(ODRec.LastName, 50); LeftJustAlphaNumeric(ODRec.Suffix, 10); ODRec.Gender; ODRec.DateOfBirth;
    Print #RptHandle, LeftJustAlphaNumeric(ODRec.AddLine1, 50); LeftJustAlphaNumeric(ODRec.AddLine2, 50); LeftJustAlphaNumeric(ODRec.City, 25);
    Print #RptHandle, LeftJustAlphaNumeric(ODRec.State, 2); LeftJustNumeric(ODRec.Zip, 9, -1); LeftJustAlphaNumeric(ODRec.OutOfCntryAdd, 50); 'changed zip from right to left fill on 10/20/08
    If QPTrim$(ODRec.EligibleDate) = "0" Then
      ODRec.EligibleDate = ""
    End If
    Print #RptHandle, RightJustNumeric(ODRec.JobClass, 6, -1); DeptNum; ODRec.EmployDate; ODRec.EligibleDate;
    Print #RptHandle, ODRec.PayPrdBeginDate; ODRec.PayPrdEndDate; ODRec.IncDecSalary; RightJustNumeric(Abs(ODRec.Salary), 12, 0);
    Print #RptHandle, ODRec.IncDecEmpleeCntrb; RightJustNumeric(Abs(ODRec.EmployeeCntrb), 12, 0); LeftJustAlphaNumeric(ODRec.Adjustment, 10);
    Print #RptHandle, LeftJustAlphaNumeric(ODRec.PayType, 10); RightJustNumeric(ODRec.VacHours, 5, 2); LeftJustAlphaNumeric(ODRec.ContrPdEmpPrd, 10);
    Print #RptHandle, ODRec.ContrPdEmpBegDate; ODRec.ContrPdEmpEndDate; ODRec.TerminationDate; LeftJustAlphaNumeric(ODRec.TermType, 10);
    Print #RptHandle, RightJustNumeric(ODRec.SharedPosition, 6, -1)
    TSalary = TSalary + ODRec.Salary
    TempCont = TempCont + ODRec.EmployeeCntrb
    TCnt = TCnt + 1
Deleted:
  Next x
  
  Close ODHandle
  
  OpenOrbTrailer OTHandle
  Get OTHandle, 1, OTRec
  If TSalary >= 0 Then
    OTRec.IncDecSalary = "+"
  Else
    OTRec.IncDecSalary = "-"
  End If
  If TempCont >= 0 Then
    OTRec.IncDecTtlEmpContrb = "+"
  Else
    OTRec.IncDecTtlEmpContrb = "-"
  End If
  
  Print #RptHandle, OTRec.RecType; RightJustNumeric(OTRec.AgencyNum, 5, -1); OTRec.RptPeriod; RightJustNumeric(CStr(TCnt), 11, -1); OTRec.IncDecSalary;
  Print #RptHandle, RightJustNumeric(Abs(TSalary), 12, 0); OTRec.IncDecTtlEmpContrb; RightJustNumeric(Abs(TempCont), 12, 0)
  
  Close OTHandle
  
  MsgBox ("The export file has been created successfully with the file name " & ReportFile$ & ".")
  Close
  
  OpenOrbEmpData OEHandle, NumOfOERecs
  Get OEHandle, x, OERec
 
  
  Exit Sub
  
CheckTotals:
  ReDim ODRecd(1 To 1) As Long
  
  ODCnt = 0
  TotOReg = 0
  TotReg = 0
  TotOOT = 0
  TotOT = 0
  TotOEmployeeMatch = 0
  TotOEmployerMatch = 0
  TotEmployeeMatch = 0
  TotEmployerMatch = 0
  OpenEmpData2File EHandle
  OHRec.PayPrdBeginDate = OHRec.PayPrdBeginDate
  OpenTransHistFile THandle
  For x = 1 To NumOfODRecs
    Get ODHandle, x, ODRec
    If ODRec.Deleted = True Then GoTo Deleted1
    TotOReg = TotOReg + ODRec.Salary
    TotOOT = TotOOT + ODRec.OTPay
    TotOEmployeeMatch = TotOEmployeeMatch + ODRec.EmployeeCntrb
    TotOEmployerMatch = TotOEmployerMatch + ODRec.EmployerCntrb
    For y = 1 To ODCnt
      If ODRecd(y) = ODRec.EmpRecNum Then
        Exit For
      End If
    Next y
    If y > ODCnt Then
      ODCnt = ODCnt + 1
      ReDim Preserve ODRecd(1 To ODCnt) As Long
      ODRecd(ODCnt) = ODRec.EmpRecNum
    End If
    
Deleted1:
  Next x
  
  GoSub GetTransTotal
  
  If OldRound(TotReg) <> OldRound(TotOReg) Then
    frmMessageWOpts.Label1.Caption = "The total retirement gross pay for ORBIT (" & QPTrim$(Using$("$###,###.##", TotOReg)) & ") and for stored pay transactions (" & Using$("$###,###.##", TotReg) & ") for the same pay period do not match. Continue anyway?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Continue"
    frmMessageWOpts.cmdExit.Text = "ESC Abort"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      Exit Sub
    Else
      Unload frmMessageWOpts
      MainLog ("Total ORBIT retirement gross pay (" & QPTrim$(Using$("$###,###.##", TotOReg)) & ") was not equal to retirement gross from transactions (" & Using$("$###,###.##", TotReg) & ") and was saved anyway.")
    End If
  End If
  
  If OldRound(TotOT) <> OldRound(TotOOT) Then
    frmMessageWOpts.Label1.Caption = "The total OT pay for ORBIT (" & QPTrim$(Using$("$###,###.##", TotOOT)) & ") and for stored pay transactions (" & Using$("$###,###.##", TotOT) & ") for the same pay period do not match. Continue anyway?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Continue"
    frmMessageWOpts.cmdExit.Text = "ESC Abort"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      Exit Sub
    Else
      Unload frmMessageWOpts
      MainLog ("Total ORBIT OT pay (" & QPTrim$(Using$("$###,###.##", TotOOT)) & ") was not equal to OT pay from transactions (" & Using$("$###,###.##", TotOT) & ") and was saved anyway.")
    End If
  End If
  
  If OldRound(TotEmployeeMatch) <> OldRound(TotOEmployeeMatch) Then
    frmMessageWOpts.Label1.Caption = "The total employee match amount for ORBIT (" & QPTrim$(Using$("$###,###.##", TotOEmployeeMatch)) & ") and for stored pay transactions (" & Using$("$###,###.##", TotEmployeeMatch) & ") for the same pay period do not match. Continue anyway?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Continue"
    frmMessageWOpts.cmdExit.Text = "ESC Abort"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      Exit Sub
    Else
      Unload frmMessageWOpts
      MainLog ("Total ORBIT employee match (" & QPTrim$(Using$("$###,###.##", TotOEmployeeMatch)) & ") was not equal to employee match from transactions (" & Using$("$###,###.##", TotEmployeeMatch) & ") and was saved anyway.")
    End If
  End If
  
  If OldRound(TotEmployerMatch) <> OldRound(TotOEmployerMatch) Then
    frmMessageWOpts.Label1.Caption = "The total employer match amount for ORBIT (" & QPTrim$(Using$("$###,###.##", TotOEmployerMatch)) & ") and for stored pay transactions (" & Using$("$###,###.##", TotEmployerMatch) & ") for the same pay period do not match. Continue anyway?"
    frmMessageWOpts.Label1.Top = 650
    frmMessageWOpts.cmdCont.Text = "F10 Continue"
    frmMessageWOpts.cmdExit.Text = "ESC Abort"
    frmMessageWOpts.Show vbModal
    If frmMessageWOpts.fptxtChoice.Text = "abort" Then
      Unload frmMessageWOpts
      Exit Sub
    Else
      Unload frmMessageWOpts
      MainLog ("Total ORBIT employer match (" & QPTrim$(Using$("$###,###.##", TotOEmployerMatch)) & ") was not equal to employer match from transactions (" & Using$("$###,###.##", TotEmployerMatch) & ") and was saved anyway.")
    End If
  End If
  
  Return

GetTransTotal:
  For y = 1 To ODCnt
    Get ODHandle, y, ODRec
    Get EHandle, ODRecd(y), EmpRec
    NextRec = EmpRec.LastTransRec
    EmpRec.EmpLName = EmpRec.EmpLName
    Do While NextRec > 0
      Get THandle, NextRec, TransRec
      If TransRec.CheckDate >= OHRec.PayPrdBeginDate And TransRec.CheckDate <= OHRec.PayPrdEndDate Then
        TotReg = TotReg + TransRec.RetGrossPay
        TotOT = TotOT + TransRec.TotOTWage
        TotEmployeeMatch = TotEmployeeMatch + TransRec.RetireAmt
        TotEmployerMatch = TotEmployerMatch + TransRec.MatchRetAmt
      End If
      NextRec = TransRec.PrevTransRec
    Loop
Deleted2:
  Next y
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "CreateExportFile", Erl)
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

Private Function LeftJustAlphaNumeric(ByVal Name As String, Length As Integer) As String
  Dim ThisLen As Integer
  Dim x As Integer
  On Error GoTo ERRORSTUFF
  
  Name = QPTrim$(Name)
  ThisLen = Len(Name)
  For x = ThisLen + 1 To Length
    Name = Name + " "
  Next x
  
  LeftJustAlphaNumeric = Name
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "CreateExportFile", Erl)
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
  
End Function
Private Function RightJustNumeric(ByVal Number As String, Length As Integer, NumType As Integer) As String
  Dim ThisLen As Integer
  Dim thisNum As String
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
 
  If NumType = 0 Then
    Number = QPTrim$(Using$("#########.##", CDbl(Number)))
  ElseIf NumType = 2 Then
    Number = FormatNumber(Number, 1, vbFalse)
  Else
    Number = QPTrim(Number)
  End If
  ThisLen = Len(Number)
  ThisLen = Length - ThisLen
  
  For x = 1 To ThisLen
    thisNum = thisNum + "0"
  Next x
  
  Number = thisNum & Number
  RightJustNumeric = Number
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "CreateExportFile", Erl)
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
  
End Function
Private Function LeftJustNumeric(ByVal Number As String, Length As Integer, NumType As Integer) As String
  Dim ThisLen As Integer
  Dim thisNum As String
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
 
  If NumType = 0 Then
    Number = QPTrim$(Using$("#########.##", CDbl(Number)))
  ElseIf NumType = 2 Then
    Number = FormatNumber(Number, 1, vbFalse)
  Else
    Number = QPTrim(Number)
  End If
  ThisLen = Len(Number)
  ThisLen = Length - ThisLen
  
  For x = 1 To ThisLen
    thisNum = "0" + thisNum
  Next x
  
  Number = Number & thisNum
  LeftJustNumeric = Number
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "LeftJustNumeric", Erl)
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
  
End Function


Private Function FormatThisPayPd(ByRef ThisDate As String, ByVal Vers As Integer) As String
  Dim ch As String
  Dim DateLen As Integer
  Dim FSPstn As Integer
  Dim x As Integer
  Dim ThisDay As String
  Dim ThisMonth As String
  Dim ThisYear As String
  
  On Error GoTo ERRORSTUFF
 
  FSPstn = 0
  DateLen = Len(ThisDate)
  For x = 1 To DateLen
    ch = Mid(ThisDate, x, 1)
    If ch = "/" Then
      FSPstn = x
      Exit For
    End If
  Next x
  
  ThisMonth = Mid(ThisDate, 1, FSPstn - 1)
  If Len(ThisMonth) = 1 Then ThisMonth = "0" & ThisMonth
  ThisDay = Mid(ThisDate, FSPstn + 1, 2)
  If Len(ThisDay) = 1 Then ThisDay = "0" + ThisDay
  ThisYear = Mid(ThisDate, DateLen - 3, DateLen)
  If Vers = 2 Then
    ThisDate = ThisYear & ThisMonth & ThisDay
  ElseIf Vers = 1 Then
    ThisDate = ThisYear & ThisMonth
  End If
  FormatThisPayPd = ThisDate
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "CreateExportFile", Erl)
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
 
End Function

Private Sub cmdPrint_Click()
  If Not Exist(OrbitDetail) Then
    MsgBox ("Please process ORBIT records first.")
    Exit Sub
  End If
   
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    MsgBox ("The printer pitch should be no less than pitch 12.")
    Call PrintText
    MainLog ("User printed ORBIT report in text.")
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PrintGraphics
    MainLog ("User printed ORBIT report in graphics.")
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
      SendKeys "%x"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdCreate_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%P"
      Call cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
'  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpORBITPost.")
      KillFile "orbpostprint.dat"
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmORBITMenu.Show
  KillFile "orbpostprint.dat"
  DoEvents
  Unload Me
End Sub

Private Sub PrintGraphics()
  Dim URec As UnitFileRecType
  Dim UHandle As Integer
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim OHRec As OrbitHeader
  Dim OHHandle As Integer
  Dim x As Integer
  Dim PlanCode() As String
  Dim PCPay() As Double
  Dim TotPCPay As Double
  Dim PCCnt() As Integer
  Dim PCEmpCont() As Double
  Dim TotPCEmpCont As Double
  Dim PCCityCont() As Double
  Dim TotPCCityCont As Double
  Dim TotPCCnt As Integer
  Dim JobClass() As String
  Dim JCPay() As Double
  Dim TJCPay() As Double
  Dim TJCCnt() As Double
  Dim TJCEmpCont() As Double
  Dim TJCCityCont() As Double
  Dim TotJCPay As Double
  Dim JCCnt() As Integer
  Dim JCEmpCont() As Double
  Dim TotJCEmpCont As Double
  Dim JCCityCont() As Double
  Dim TotJCCityCont As Double
  Dim TotJCCnt As Integer, t As Integer
  Dim y As Integer, z As Integer, C As Integer
  Dim ThisPlan As String
  Dim ThisJob As String
  Dim DashPos As Integer
  Dim RptName$
  Dim RptHandle As Integer
  Dim RptName2$
  Dim RptHandle2 As Integer
  Dim RptName3$
  Dim RptHandle3 As Integer
  Dim dlm As String
  Dim Town As String
  Dim ThisJobClass As String
  Dim Test As Integer
  Dim PrintPay As String
  Dim PrintEmpCont As String
  Dim PrintCityCont As String
  Dim CntWODel As Integer
  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenUnitFile UHandle
  Get UHandle, 1, URec
  Close UHandle
  Town = QPTrim$(URec.UFCITY)
  OpenOrbHeader OHHandle
  Get OHHandle, 1, OHRec
  Close OHHandle
  
  dlm = "~"
  ReDim PlanCode(1 To 19) As String
  ReDim PCPay(1 To 19) As Double
  ReDim PCCnt(1 To 19) As Integer
  ReDim PCEmpCont(1 To 19) As Double
  ReDim PCCityCont(1 To 19) As Double
  ReDim PrintLine1(1 To 28) As String
  ReDim PrintLine2(1 To 28) As String
  ReDim PrintLine3(1 To 28) As String
  ReDim PrintLine4(1 To 28) As String
  
  PlanCode(1) = "STG - Teachers' and State Employees' General Class"
  PlanCode(2) = "STL - Teachers' and State Employees' Law Enforcement Class"
  PlanCode(3) = "STMAX - Teachers' and State Employees' earnings exceeding IRS max"
  PlanCode(4) = "STRS -  - Teachers' and State Employees' rehired retirees subject to earnings limits"
  PlanCode(5) = "STRE - Teachers' and State Employees' rehired retirees exempt from earnings limits "
  PlanCode(6) = "STDIS - Teachers' and State Employees' on Short Term Disability"
  PlanCode(7) = "LOCG - Local Governmental Employees System General Class"
  PlanCode(8) = "LOCL - Local Governmental Employees System Law Enforcement Class"
  PlanCode(9) = "LOCF - Local Governmental Employees System Firemen Class "
  PlanCode(10) = "LOCMAX - Local Governmental Employees System earnings exceeding IRS max"
  PlanCode(11) = "LOCWP - Local Governmental Employees System probationary employees"
  PlanCode(12) = "LOCROD - Local Governmental Employees System Register Of Deeds Class"
  PlanCode(13) = "LOCRS - Local Governmental Employees System rehired retirees subject to earnings limits"
  PlanCode(14) = "JUD1 - Consolidated Judicial System = Justices of the Supreme Court and Judges of the Court of Appeals"
  PlanCode(15) = "JUD2 - Consolidated Judicial System = Judges of the Superior Court and Administrative Officer of the Courts"
  PlanCode(16) = "JUD3 - Consolidated Judicial System = Judges of the District Court, District Attorney, Clerk of Superior Court"
  PlanCode(17) = "LEGL - Legislative Retirement System Legislators"
  PlanCode(18) = "ORPG - Optional Retirement Program for applicable University and Community College personnel"
  PlanCode(19) = "ORPMAX - Optional Retirement Program for employees exceeding IRS max"
  
  ReDim JobClass(1 To 19, 1 To 28) As String
  For x = 1 To 19
    JobClass(x, 1) = "100 Educational Professionals"
    JobClass(x, 2) = "102 Educational Executives"
    JobClass(x, 3) = "103 Educational Management"
    JobClass(x, 4) = "104 Educational Administrative Support Personnel"
    JobClass(x, 5) = "105 Educational Support Personnel"
    JobClass(x, 6) = "200 Professionals"
    JobClass(x, 7) = "201 Government Officials and Executives"
    JobClass(x, 8) = "202 Administrative"
    JobClass(x, 9) = "203 Technical and Trades"
    JobClass(x, 10) = "204 Skilled Labor"
    JobClass(x, 11) = "300 University and Community College Professionals"
    JobClass(x, 12) = "301 University and Community College Executives"
    JobClass(x, 13) = "302 University and Community College Management"
    JobClass(x, 14) = "303 University Agricultural (AG) Extension"
    JobClass(x, 15) = "400 Local Firefighters"
    JobClass(x, 16) = "401 Local Emergency workers"
    JobClass(x, 17) = "500 SHP Trooper"
    JobClass(x, 18) = "501 Police Officer"
    JobClass(x, 19) = "502 Public Safety Officer"
    JobClass(x, 20) = "503 Sheriff"
    JobClass(x, 21) = "504 Deputy Sheriff"
    JobClass(x, 22) = "505 Wildlife Officer"
    JobClass(x, 23) = "506 Revenue Enforcement Officer"
    JobClass(x, 24) = "507 ALE Officer"
    JobClass(x, 25) = "508 Jailer (Certified)"
    JobClass(x, 26) = "509 Protective Services (Non-Certified)"
    JobClass(x, 27) = "600 Health Care Professionals"
    JobClass(x, 28) = "601 Social Service Professionals"
  Next x
  ReDim JCPay(1 To 19, 1 To 28) As Double
  ReDim JCCnt(1 To 19, 1 To 28) As Integer
  ReDim JCEmpCont(1 To 19, 1 To 28) As Double
  ReDim JCCityCont(1 To 19, 1 To 28) As Double
  
  ReDim TJCPay(1 To 28) As Double
  ReDim TJCCnt(1 To 28) As Double
  ReDim TJCEmpCont(1 To 28) As Double
  ReDim TJCCityCont(1 To 28) As Double
  
  Call NameSort
  RptName$ = "PRRPTS\ORBIT.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  RptName2$ = "PRRPTS\ORBITSUB.RPT"
  RptHandle2 = FreeFile
  Open RptName2$ For Output As RptHandle2
  RptName3$ = "PRRPTS\ORBITSUB2.RPT"
  RptHandle3 = FreeFile
  Open RptName3$ For Output As RptHandle3
  
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  
  CntWODel = 0
  OpenOrbDetail OHandle, NumOfORecs
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo NotThis
    CntWODel = CntWODel + 1
NotThis:
  Next x
  
  For y = 1 To 19
    DashPos = InStr(PlanCode(y), "-")
    ThisPlan = QPTrim$(Mid(PlanCode(y), 1, DashPos - 1))
    For x = 1 To NumOfORecs 'CntWODel
    Get OHandle, NameIdx(x), ORec
    If ORec.Deleted = True Then GoTo Deleted
    If QPTrim$(ORec.PlanCode) = ThisPlan Then
      PCPay(y) = PCPay(y) + ORec.Salary
      TotPCPay = TotPCPay + ORec.Salary
      PCCnt(y) = PCCnt(y) + 1
      TotPCCnt = TotPCCnt + 1
      PCEmpCont(y) = PCEmpCont(y) + ORec.EmployeeCntrb
      TotPCEmpCont = TotPCEmpCont + ORec.EmployeeCntrb
      PCCityCont(y) = PCCityCont(y) + ORec.EmployerCntrb
      TotPCCityCont = TotPCCityCont + ORec.EmployerCntrb
      For z = 1 To 28
        ThisJob = Mid(JobClass(y, z), 1, 3)
        If QPTrim$(ORec.JobClass) = ThisJob Then
          JCPay(y, z) = JCPay(y, z) + ORec.Salary
          TJCPay(z) = TJCPay(z) + ORec.Salary
          TotJCPay = TotJCPay + ORec.Salary
          JCCnt(y, z) = JCCnt(y, z) + 1
          TJCCnt(z) = TJCCnt(z) + 1
          TotJCCnt = TotJCCnt + 1
          JCEmpCont(y, z) = JCEmpCont(y, z) + ORec.EmployeeCntrb
          TJCEmpCont(z) = TJCEmpCont(z) + ORec.EmployeeCntrb
          TotJCEmpCont = TotJCEmpCont + ORec.EmployeeCntrb
          JCCityCont(y, z) = JCCityCont(y, z) + ORec.EmployerCntrb
          TJCCityCont(z) = TJCCityCont(z) + ORec.EmployerCntrb
          TotJCCityCont = TotJCCityCont + ORec.EmployerCntrb
          ThisJobClass = EmptyFill(JobClass(y, z), 75)
          PrintPay = EmptyFill(Using$("$##,###.##", JCPay(y, z)), 15)
          PrintEmpCont = EmptyFill(Using$("$##,###.##", JCEmpCont(y, z)), 15)
          PrintCityCont = EmptyFill(Using$("$##,###.##", JCCityCont(y, z)), 15)
          PrintLine1(z) = ThisJobClass
          PrintLine2(z) = PrintPay
          PrintLine3(z) = PrintEmpCont
          PrintLine4(z) = PrintCityCont
        End If
      Next z
      GoSub PrintIt
     End If
Deleted:
   Next x
   For t = 1 To 28
     PrintLine1(t) = ""
     PrintLine2(t) = ""
     PrintLine3(t) = ""
     PrintLine4(t) = ""
   Next t
   FrmShowPctComp.ShowPctComp y, 19
   If FrmShowPctComp.Out = True Then
     Close
     FrmShowPctComp.Out = False
     EnableCloseButton Me.hwnd, True
     Me.cmdExit.Enabled = True
     Unload FrmShowPctComp
     Exit Sub
   End If
  Next y
  GoSub PrintItSub
  GoSub PrintItSub2
  
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Unload FrmShowPctComp
 
  Close
  
  arORBIT.Show
  
  AHandle = FreeFile
  Open "orbpostprint.dat" For Output As AHandle
  Print #AHandle, "Done"
  Close AHandle
  
  Exit Sub
  
PrintIt:
  '                   0                1                        2                      3
  Print #RptHandle, Town; dlm; OHRec.AgencyNum; dlm; OHRec.FileCreateDate; dlm; OHRec.FrmtVersion; dlm;
  '                          4                           5                        6
  Print #RptHandle, OHRec.PayPrdBeginDate; dlm; OHRec.PayPrdEndDate; dlm; OHRec.RptPeriod; dlm;
  '                       7                   8                      9                  10
  Print #RptHandle, ORec.LastName; dlm; ORec.FirstName; dlm; ORec.MiddleName; dlm; ORec.Suffix; dlm;
  '                        11                        12                    13                  14
  Print #RptHandle, ORec.EmployeeCntrb; dlm; ORec.EmployerCntrb; dlm; ORec.JobClass; dlm; ORec.MemberID; dlm;
  '                     15                   16                         17                    18
  Print #RptHandle, ORec.OTPay; dlm; ORec.PayPrdBeginDate; dlm; ORec.PayPrdEndDate; dlm; PlanCode(y); dlm;
  '                     19                20
  Print #RptHandle, ORec.Salary; dlm; ORec.EmpNum; dlm;
  
  For t = 1 To 28
    Print #RptHandle, PrintLine1(t); dlm;
    Print #RptHandle, PrintLine2(t); dlm;
    Print #RptHandle, PrintLine3(t); dlm;
    Print #RptHandle, PrintLine4(t); dlm;
  Next t
  '
  Print #RptHandle, TotJCPay; dlm; TotJCCnt; dlm; TotJCEmpCont; dlm; TotJCCityCont; dlm;
  '
  Print #RptHandle, TotPCPay; dlm; TotPCCnt; dlm; TotPCEmpCont; dlm; TotPCCityCont; dlm; ORec.PayType; dlm; ORec.Adjustment
  
  Return
  
PrintItSub:
  For x = 1 To 28
    If TJCCnt(x) > 0 Then
      Print #RptHandle2, JobClass(1, x); dlm; TJCCnt(x); dlm; TJCPay(x); dlm; TJCEmpCont(x); dlm; TJCCityCont(x)
    End If
  Next x
  
  Return
  
PrintItSub2:
  For x = 1 To 19
    If PCCnt(x) > 0 Then
      Print #RptHandle3, PlanCode(x); dlm; PCCnt(x); dlm; PCPay(x); dlm; PCEmpCont(x); dlm; PCCityCont(x)
    End If
  Next x
  
  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITPost", "PrintGraphics", Erl)
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
Private Sub PrintText()
  Dim URec As UnitFileRecType
  Dim UHandle As Integer
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim OHRec As OrbitHeader
  Dim OHHandle As Integer
  Dim x As Integer
  Dim PlanCode() As String
  Dim PCPay() As Double
  Dim TotPCPay As Double
  Dim TotPCOTPay As Double
  Dim PCCnt() As Integer
  Dim PCEmpCont() As Double
  Dim TotPCEmpCont As Double
  Dim PCCityCont() As Double
  Dim TotPCCityCont As Double
  Dim TotPCCnt As Integer
  Dim JobClass() As String
  Dim JCPay() As Double
  Dim TJCPay() As Double
  Dim TJCCnt() As Double
  Dim TJCEmpCont() As Double
  Dim TJCCityCont() As Double
  Dim TotJCPay As Double
  Dim JCCnt() As Integer
  Dim JCEmpCont() As Double
  Dim TotJCEmpCont As Double
  Dim JCCityCont() As Double
  Dim TotJCCityCont As Double
  Dim TotJCCnt As Integer, t As Integer
  Dim y As Integer, z As Integer, C As Integer
  Dim ThisPlan As String
  Dim ThisJob As String
  Dim DashPos As Integer
  Dim RptName$, RptTitle$
  Dim RptHandle As Integer
  Dim RptName2$
  Dim RptHandle2 As Integer
  Dim RptName3$
  Dim RptHandle3 As Integer
  Dim dlm As String
  Dim Town As String
  Dim ThisJobClass As String
  Dim Test As Integer
  Dim PrintPay As String
  Dim PrintEmpCont As String
  Dim PrintCityCont As String
  Dim FF$, Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim Dash As String * 80
  Dim EmpNo As String * 10
  Dim EmpFName As String
  Dim EmpLName As String
  Dim FullName As String * 48
  Dim MemID As String * 10
  Dim ThisJC As String * 25
  Dim ThisPC As String * 25
  Dim CntWODel As Integer
  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
  FF$ = Chr(12)
  MaxLines = 57
  LineCnt = 0
  
  OpenUnitFile UHandle
  Get UHandle, 1, URec
  Close UHandle
  Town = QPTrim$(URec.UFCITY)
  OpenOrbHeader OHHandle
  Get OHHandle, 1, OHRec
  Close OHHandle
  
  dlm = "~"
  ReDim PlanCode(1 To 19) As String
  ReDim PCPay(1 To 19) As Double
  ReDim PCCnt(1 To 19) As Integer
  ReDim PCOTPay(1 To 19) As Double
  ReDim PCEmpCont(1 To 19) As Double
  ReDim PCCityCont(1 To 19) As Double
  ReDim PrintLine1(1 To 28) As String
  ReDim PrintLine2(1 To 28) As String
  ReDim PrintLine3(1 To 28) As String
  ReDim PrintLine4(1 To 28) As String
  
  PlanCode(1) = "STG - Teachers' and State Employees' General Class"
  PlanCode(2) = "STL - Teachers' and State Employees' Law Enforcement Class"
  PlanCode(3) = "STMAX - Teachers' and State Employees' earnings exceeding IRS max"
  PlanCode(4) = "STRS -  - Teachers' and State Employees' rehired retirees subject to earnings limits"
  PlanCode(5) = "STRE - Teachers' and State Employees' rehired retirees exempt from earnings limits "
  PlanCode(6) = "STDIS - Teachers' and State Employees' on Short Term Disability"
  PlanCode(7) = "LOCG - Local Governmental Employees System General Class"
  PlanCode(8) = "LOCL - Local Governmental Employees System Law Enforcement Class"
  PlanCode(9) = "LOCF - Local Governmental Employees System Firemen Class "
  PlanCode(10) = "LOCMAX - Local Governmental Employees System earnings exceeding IRS max"
  PlanCode(11) = "LOCWP - Local Governmental Employees System probationary employees"
  PlanCode(12) = "LOCROD - Local Governmental Employees System Register Of Deeds Class"
  PlanCode(13) = "LOCRS - Local Governmental Employees System rehired retirees subject to earnings limits"
  PlanCode(14) = "JUD1 - Consolidated Judicial System = Justices of the Supreme Court and Judges of the Court of Appeals"
  PlanCode(15) = "JUD2 - Consolidated Judicial System = Judges of the Superior Court and Administrative Officer of the Courts"
  PlanCode(16) = "JUD3 - Consolidated Judicial System = Judges of the District Court, District Attorney, Clerk of Superior Court"
  PlanCode(17) = "LEGL - Legislative Retirement System Legislators"
  PlanCode(18) = "ORPG - Optional Retirement Program for applicable University and Community College personnel"
  PlanCode(19) = "ORPMAX - Optional Retirement Program for employees exceeding IRS max"
  
  ReDim JobClass(1 To 19, 1 To 28) As String
  For x = 1 To 19
    JobClass(x, 1) = "100 Educational Professionals"
    JobClass(x, 2) = "102 Educational Executives"
    JobClass(x, 3) = "103 Educational Management"
    JobClass(x, 4) = "104 Educational Administrative Support Personnel"
    JobClass(x, 5) = "105 Educational Support Personnel"
    JobClass(x, 6) = "200 Professionals"
    JobClass(x, 7) = "201 Government Officials and Executives"
    JobClass(x, 8) = "202 Administrative"
    JobClass(x, 9) = "203 Technical and Trades"
    JobClass(x, 10) = "204 Skilled Labor"
    JobClass(x, 11) = "300 University and Community College Professionals"
    JobClass(x, 12) = "301 University and Community College Executives"
    JobClass(x, 13) = "302 University and Community College Management"
    JobClass(x, 14) = "303 University Agricultural (AG) Extension"
    JobClass(x, 15) = "400 Local Firefighters"
    JobClass(x, 16) = "401 Local Emergency workers"
    JobClass(x, 17) = "500 SHP Trooper"
    JobClass(x, 18) = "501 Police Officer"
    JobClass(x, 19) = "502 Public Safety Officer"
    JobClass(x, 20) = "503 Sheriff"
    JobClass(x, 21) = "504 Deputy Sheriff"
    JobClass(x, 22) = "505 Wildlife Officer"
    JobClass(x, 23) = "506 Revenue Enforcement Officer"
    JobClass(x, 24) = "507 ALE Officer"
    JobClass(x, 25) = "508 Jailer (Certified)"
    JobClass(x, 26) = "509 Protective Services (Non-Certified)"
    JobClass(x, 27) = "600 Health Care Professionals"
    JobClass(x, 28) = "601 Social Service Professionals"
  Next x
  ReDim JCPay(1 To 19, 1 To 28) As Double
  ReDim JCCnt(1 To 19, 1 To 28) As Integer
  ReDim JCEmpCont(1 To 19, 1 To 28) As Double
  ReDim JCCityCont(1 To 19, 1 To 28) As Double
  
  ReDim TJCPay(1 To 28) As Double
  ReDim TJCCnt(1 To 28) As Double
  ReDim TJCEmpCont(1 To 28) As Double
  ReDim TJCCityCont(1 To 28) As Double
  
  Call NameSort
  RptTitle$ = "NC ORBIT Report"
  RptName$ = "PRRPTS\TxtORBIT.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  GoSub PrintHeader
  
  ReDim HasPay(1 To 19) As Boolean
  For x = 1 To 19
     HasPay(x) = False
  Next x
  
'  CntWODel = 0
  OpenOrbDetail OHandle, NumOfORecs
'  For x = 1 To NumOfORecs
'    Get OHandle, x, ORec
'    If ORec.Deleted = True Then GoTo NotThis
'    CntWODel = CntWODel + 1
'NotThis:
'  Next x
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  
  For x = 1 To NumOfORecs 'CntWODel
    Get OHandle, x, ORec
    If ORec.Deleted = True Then GoTo SkipDeleted
    If ORec.Salary > 0 Then
      For t = 1 To 19
        DashPos = InStr(PlanCode(t), "-")
        ThisPlan = QPTrim$(Mid(PlanCode(t), 1, DashPos - 1))
        If QPTrim$(ORec.PlanCode) = ThisPlan Then
          HasPay(t) = True
          Exit For
        End If
      Next t
    End If
SkipDeleted:
  Next x
  
  For y = 1 To 19
    If HasPay(y) = False Then GoTo Deleted
    DashPos = InStr(PlanCode(y), "-")
    ThisPlan = QPTrim$(Mid(PlanCode(y), 1, DashPos - 1))
    GoSub PrintPayHeader
    For x = 1 To NumOfORecs 'CntWODel
    Get OHandle, NameIdx(x), ORec
    If ORec.Deleted = True Then GoTo Deletedx
    If QPTrim$(ORec.PlanCode) = ThisPlan Then
      PCPay(y) = PCPay(y) + ORec.Salary
      TotPCPay = TotPCPay + ORec.Salary
      PCOTPay(y) = PCOTPay(y) + ORec.OTPay
      TotPCOTPay = TotPCOTPay + ORec.OTPay
      PCCnt(y) = PCCnt(y) + 1
      TotPCCnt = TotPCCnt + 1
      PCEmpCont(y) = PCEmpCont(y) + ORec.EmployeeCntrb
      TotPCEmpCont = TotPCEmpCont + ORec.EmployeeCntrb
      PCCityCont(y) = PCCityCont(y) + ORec.EmployerCntrb
      TotPCCityCont = TotPCCityCont + ORec.EmployerCntrb
      For z = 1 To 28
        ThisJob = Mid(JobClass(y, z), 1, 3)
        If QPTrim$(ORec.JobClass) = ThisJob Then
          JCPay(y, z) = JCPay(y, z) + ORec.Salary
          TJCPay(z) = TJCPay(z) + ORec.Salary
          TotJCPay = TotJCPay + ORec.Salary
          JCCnt(y, z) = JCCnt(y, z) + 1
          TJCCnt(z) = TJCCnt(z) + 1
          TotJCCnt = TotJCCnt + 1
          JCEmpCont(y, z) = JCEmpCont(y, z) + ORec.EmployeeCntrb
          TJCEmpCont(z) = TJCEmpCont(z) + ORec.EmployeeCntrb
          TotJCEmpCont = TotJCEmpCont + ORec.EmployeeCntrb
          JCCityCont(y, z) = JCCityCont(y, z) + ORec.EmployerCntrb
          TJCCityCont(z) = TJCCityCont(z) + ORec.EmployerCntrb
          TotJCCityCont = TotJCCityCont + ORec.EmployerCntrb
        End If
      Next z
      GoSub PrintPayCode
     End If
Deletedx:
   Next x
   GoSub PrintPayFooter
Deleted:
   FrmShowPctComp.ShowPctComp y, 19
   If FrmShowPctComp.Out = True Then
     Close
     FrmShowPctComp.Out = False
     EnableCloseButton Me.hwnd, True
     Me.cmdExit.Enabled = True
     Unload FrmShowPctComp
     Exit Sub
   End If
  Next y
  
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Unload FrmShowPctComp
  
  GoSub PrintRptFooter
  Print #RptHandle, FF$
  
  Close
  
  ViewPrint RptName$, RptTitle$
  AHandle = FreeFile
  Open "orbpostprint.dat" For Output As AHandle
  Print #AHandle, "Done"
  Close AHandle
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(35); "NC ORBIT Report"
  Print #RptHandle,
  Print #RptHandle, "Employer: " + QPTrim$(URec.UFEMPR); Tab(81); "Page# " + CStr(Page)
  Print #RptHandle, "Agency #: " & QPTrim$(OHRec.AgencyNum)
  Print #RptHandle, "Reporting Period: " & MakeRptDate(OHRec.RptPeriod, 1); Tab(58); "File Creation Date: " & MakeRegDate(OHRec.FileCreateDate)
  Print #RptHandle, "Report Print Date: " + CStr(Date); Tab(60); "Pay Period Start: " & MakeRegDate(OHRec.PayPrdBeginDate)
  Print #RptHandle, "Format Version: " & OHRec.FrmtVersion; Tab(62); "Pay Period End: " & MakeRegDate(OHRec.PayPrdEndDate)
  Print #RptHandle, String$(87, "-")
  Print #RptHandle,
  LineCnt = 9
  
  Return
  
PrintPayHeader:
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, PlanCode(y)
  Print #RptHandle, "Emp PR#"; Tab(14); "Emp Name"; Tab(66); "Pay Amt"; Tab(82); "OT Amt"
  Print #RptHandle, "Emp Member#"; Tab(14); "Job Class"; Tab(25); "Pay Pd Start"; Tab(39); "Pay Pd End"; Tab(60); "Emplyee Contr"; Tab(75); "Emplyer Contr"
  Print #RptHandle, "Pay Type"; Tab(14); "Adjustment Type"
  Print #RptHandle, String$(87, "-")
  LineCnt = LineCnt + 5
  
  Return
  
PrintPayCode:
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintPayHeader
  End If
  EmpLName = QPTrim$(ORec.LastName)
  EmpFName = QPTrim$(ORec.FirstName)
  LSet FullName = EmpLName & ", " & EmpFName
  RSet MemID = QPTrim$(ORec.MemberID)
  RSet EmpNo = QPTrim$(ORec.EmpNum)
  Print #RptHandle, EmpNo; Tab(14); FullName; Tab(63); Using$("$##,###.##", ORec.Salary); Tab(78); Using$("$##,###.##", ORec.OTPay)
  Print #RptHandle, MemID; Tab(14); QPTrim$(ORec.JobClass); Tab(26); MakeRptDate(ORec.PayPrdBeginDate, 2); Tab(39); MakeRptDate(ORec.PayPrdEndDate, 2); Tab(63); Using$("$##,###.##", ORec.EmployeeCntrb); Tab(78); Using$("$##,###.##", ORec.EmployerCntrb)
  Print #RptHandle, QPTrim$(ORec.PayType); Tab(14); QPTrim$(ORec.Adjustment)
  Print #RptHandle,
  LineCnt = LineCnt + 4
    
  Return
  
PrintPayFooter:
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, String$(87, "-")
  Print #RptHandle, Tab(26); "# Trans"; Tab(37); "Total Pay"; Tab(50); "Ttl Emplyee Contr"; Tab(71); "Ttl Emplyer Contr"
  Print #RptHandle, "Plan Code Totals: "; Tab(27); Using$("#,###", PCCnt(y)); Tab(36); Using$("$##,###.##", PCPay(y)); Tab(57); Using$("$##,###.##", PCEmpCont(y)); Tab(78); Using$("$##,###.##", PCCityCont(y))
  Print #RptHandle, String$(87, "-")
  Print #RptHandle, "Job Class Totals: "
  LineCnt = LineCnt + 4
  For t = 1 To 28
    If JCPay(y, t) > 0 Then
      LSet ThisJC = JobClass(y, t)
      Print #RptHandle, ThisJC; Tab(27); Using$("#,###", JCCnt(y, t)); Tab(36); Using$("$##,###.##", JCPay(y, t)); Tab(57); Using$("$##,###.##", JCEmpCont(y, t)); Tab(78); Using$("$##,###.##", JCCityCont(y, t))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines - 5 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        Print #RptHandle, String$(87, "-")
        Print #RptHandle, "Job Class Totals: "; Tab(27); "# Trans"; Tab(37); "Total Pay"; Tab(50); "Ttl Emplyee Contr"; Tab(71); "Ttl Emplyer Contr"
        LineCnt = LineCnt + 2
      End If
    End If
  Next t
  For t = y + 1 To 19
    If HasPay(t) = True Then
      Print #RptHandle,
      Print #RptHandle, String$(87, "-")
      Print #RptHandle, String$(87, "-")
      LineCnt = LineCnt + 3
      Exit For
    End If
  Next t
    
  Return

PrintRptFooter:
  Print #RptHandle, FF$
  Page = Page + 1
  Print #RptHandle, Tab(31); "NC ORBIT Report Summary"
  Print #RptHandle,
  Print #RptHandle, "Employer: " + QPTrim$(URec.UFEMPR); Tab(81); "Page# " + CStr(Page)
  Print #RptHandle, "Agency #: " & QPTrim$(OHRec.AgencyNum)
  Print #RptHandle, "Reporting Period: " & MakeRptDate(OHRec.RptPeriod, 1); Tab(58); "File Creation Date: " & MakeRegDate(OHRec.FileCreateDate)
  Print #RptHandle, "Report Print Date: " + CStr(Date); Tab(60); "Pay Period Start: " & MakeRegDate(OHRec.PayPrdBeginDate)
  Print #RptHandle, "Format Version: " & OHRec.FrmtVersion; Tab(62); "Pay Period End: " & MakeRegDate(OHRec.PayPrdEndDate)
  Print #RptHandle, String$(87, "-")
  Print #RptHandle,
  Print #RptHandle, "Plan Code Totals: "
  Print #RptHandle, Tab(26); "# Trans"; Tab(37); "Total Pay"; Tab(50); "Ttl Emplyee Contr"; Tab(71); "Ttl Emplyer Contr"
  For y = 1 To 19
    If PCCnt(y) > 0 Then
      LSet ThisPC = PlanCode(y)
      Print #RptHandle, ThisPC; Tab(27); Using$("#,###", PCCnt(y)); Tab(36); Using$("$##,###.##", PCPay(y)); Tab(57); Using$("$##,###.##", PCEmpCont(y)); Tab(78); Using$("$##,###.##", PCCityCont(y))
    End If
  Next y
  Print #RptHandle, String$(87, "-")
  Print #RptHandle, "Plan Code Grand Totals "; Tab(27); Using$("#,###", TotPCCnt); Tab(36); Using$("$##,###.##", TotPCPay); Tab(57); Using$("$##,###.##", TotPCEmpCont); Tab(78); Using$("$##,###.##", TotPCCityCont)
  Print #RptHandle, String$(87, "-")
  Print #RptHandle, "Job Class Totals: "
  Print #RptHandle,
  For t = 1 To 28
    If TJCCnt(t) > 0 Then
      LSet ThisJC = JobClass(1, t)
      Print #RptHandle, ThisJC; Tab(27); Using$("#,###", TJCCnt(t)); Tab(36); Using$("$##,###.##", TJCPay(t)); Tab(57); Using$("$##,###.##", TJCEmpCont(t)); Tab(78); Using$("$##,###.##", TJCCityCont(t))
    End If
  Next t
  Print #RptHandle, String$(87, "-")
  Print #RptHandle, "Job Class Grand Totals "; Tab(27); Using$("#,###", TotJCCnt); Tab(36); Using$("$##,###.##", TotJCPay); Tab(57); Using$("$##,###.##", TotJCEmpCont); Tab(78); Using$("$##,###.##", TotJCCityCont)
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITPost", "PrintText", Erl)
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
Private Sub NameSort()
  Dim ORec As OrbitDetail
  Dim OHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  Dim Big As String
  Dim Thisx As Integer
  Dim BigSave As String
  Dim NextRec As Integer
  Dim SaveName As String
  Dim SaveNum As String
  Dim TempCnt As Integer
  Dim SaveBig As String
  Dim ThisRec As Integer
  Dim SaveRec As Integer
  Dim CheckDate As String
  Dim BegDate As String
  Dim EndDate As String
  Dim PayPeriod As String
  
  On Error GoTo ERRORSTUFF
  
  OpenOrbDetail OHandle, NumOfORecs
  ReDim TempOIDX(1 To 1) As String
  ReDim TempNameIDX(1 To 1) As String
  ReDim TempIdx(1 To 1) As Integer
  TempCnt = 0
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
'    If ORec.Deleted = True Then GoTo Nope
    TempCnt = TempCnt + 1
    ReDim Preserve TempOIDX(1 To TempCnt) As String
    ReDim Preserve TempNameIDX(1 To TempCnt) As String
    ReDim Preserve TempIdx(1 To TempCnt) As Integer
    BegDate = Mid(ORec.PayPrdBeginDate, 5, 2) & "/" & Mid(ORec.PayPrdBeginDate, 7, 2)
    EndDate = Mid(ORec.PayPrdEndDate, 5, 2) & "/" & Mid(ORec.PayPrdEndDate, 7, 2)
    PayPeriod = BegDate & " - " & EndDate
    
    CheckDate = MakeRegDate(ORec.CheckDate)
    TempOIDX(TempCnt) = PayPeriod & " " & QPTrim(ORec.EmpNum) & " " & QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName)
    TempNameIDX(TempCnt) = QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName)
    TempIdx(TempCnt) = x
Nope:
  Next x
  
  Big = ""
  For x = 1 To NumOfORecs
    Get OHandle, x, ORec
'    If ORec.Deleted = True Then GoTo NoNo
    If ORec.LastName > Big Then
      Big = QPTrim$(ORec.LastName) & ", " & QPTrim$(ORec.FirstName)
    End If
NoNo:
  Next x
  Close OHandle
  SaveBig = Big + "z"
  
  Big = SaveBig
  NextRec = 1
  Do
    For x = NextRec To TempCnt
      If TempNameIDX(x) < Big Then
        Big = TempNameIDX(x)
        ThisRec = x
      End If
    Next x
    SaveName = TempNameIDX(NextRec)
    SaveNum = TempOIDX(NextRec)
    SaveRec = TempIdx(NextRec)
    TempNameIDX(NextRec) = TempNameIDX(ThisRec)
    TempOIDX(NextRec) = TempOIDX(ThisRec)
    TempIdx(NextRec) = TempIdx(ThisRec)
    TempNameIDX(ThisRec) = SaveName
    TempOIDX(ThisRec) = SaveNum
    TempIdx(ThisRec) = SaveRec
    NextRec = NextRec + 1
    If NextRec > TempCnt Then Exit Do
    Big = SaveBig
  Loop
  NameCnt = TempCnt
  ReDim NameIdx(1 To NameCnt) As Integer
  For x = 1 To NameCnt
     NameIdx(x) = TempIdx(x)
  Next x
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "NameSort", Erl)
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

Public Function EmptyFill(ByRef ThisDesc$, ThisLen As Integer) As String
  Dim x As Integer
  Dim DescLen As Integer
  
  DescLen = Len(ThisDesc)
  For x = DescLen + 1 To ThisLen
    ThisDesc = ThisDesc + " "
  Next x
  EmptyFill = ThisDesc
End Function

Private Function MakeRptDate(ByVal ThisDate As String, ThisType As Integer) As String
 Dim ThisMn As String
 Dim ThisYr As String
 Dim ThisDay As String
 
 On Error GoTo ERRORSTUFF
 
 ThisMn = Mid(ThisDate, 5, 2)
 ThisYr = Mid(ThisDate, 1, 4)
 If ThisType = 2 Then GoTo Type2
 
 ThisDate = ThisMn & "/" & ThisYr
 MakeRptDate = ThisDate
 
 Exit Function
 
Type2:
 ThisDay = Mid(ThisDate, 7, 2)
 ThisDate = ThisMn & "/" & ThisDay & "/" & ThisYr
 MakeRptDate = ThisDate
 Exit Function
 
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmORBITMenu", "MakeRptDate", Erl)
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
  

End Function
