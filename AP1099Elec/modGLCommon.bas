Attribute VB_Name = "modGLCommon"
Option Explicit
Dim GLSetup    As GLSetupRecType
Dim GLFund     As GLFundRecType
Dim GLAcct     As GLAcctRecType
Dim GLFundIdx  As GLFundIndexType
Dim GLAcctidx  As GLAcctIndexType
Dim GLDept     As GLDeptRecType
Dim GLDeptIdx  As GLDeptIndexType
Dim GLBank     As GLBankRecType
Dim APInvTax   As APInvTaxRecType
Dim GJEdit     As TrEditRecType
Dim GLTrans    As GLTransRecType
Dim CJEdit     As CJEditRecType
Dim BgtEdit    As TrEditRecType
Dim BgtTrans   As GLTransRecType
Dim OSChek     As OSChekRecType
Dim ApLedger   As APLedger81RecType
Dim APDist     As APDistRecType
Dim apvendor   As VendorRecType
Dim GLFNCT     As GLFNCTRecType
Dim GLFNCTIdx  As GLFNCTIndexType

Public screenW As Long
Public coladj As Double
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
'Put these 2 with procedure
'Dim RetValue As Integer
'  RetValue = sndPlaySound("UBToil.dat", SND_ASYNC Or SND_NODEFAULT)
Global StartPath As String
Global PrYr As Integer
Global rptopt As Integer
Global zoomindex As Integer

Public Sub ActivateControls(fmx As Form, Optional op As Boolean)
  Dim x As Control, cnt As Integer
      For cnt = 0 To fmx.Count - 1
      Set x = fmx.Controls.Item(cnt)
        If TypeOf x Is CommandButton Then
          x.Enabled = True
        End If
        If TypeOf x Is fpCombo Then
          x.Enabled = True
        End If
        If TypeOf x Is fpDateTime Then
          x.Enabled = True
        End If
        If TypeOf x Is fpMask Then
          x.Enabled = True
        End If
        If TypeOf x Is TextBox Then
          x.Enabled = True
        End If
      Next cnt
      If op = True Then
        fmx.mnuOptions.Enabled = True
      End If
     EnableCloseButton fmx.hwnd, True
     Screen.MousePointer = vbDefault
End Sub
Public Sub DeActivateControls(fmx As Form, Optional op As Boolean)
   Dim cnt As Integer, x As Control
      For cnt = 0 To fmx.Count - 1
      Set x = fmx.Controls.Item(cnt)
        If TypeOf x Is CommandButton Then
          x.Enabled = False
        End If
        If TypeOf x Is fpCombo Then
          x.Enabled = False
        End If
        If TypeOf x Is fpDateTime Then
          x.Enabled = False
        End If
        If TypeOf x Is fpMask Then
          x.Enabled = False
        End If
        If TypeOf x Is TextBox Then
          x.Enabled = False
        End If
        If TypeOf x Is Menu Then
          x.Enabled = True
        End If
      Next cnt
      If op = True Then
        fmx.mnuOptions.Enabled = False
      End If
     EnableCloseButton fmx.hwnd, False
     Screen.MousePointer = vbHourglass
End Sub

Public Function FindAcct(FundNum$, GLFundLen)
'To search for Account - see if Fund Has been used so can delete fund or not
  Dim NumOfAccts As Integer, CntA As Integer, AcctFile As Integer
  Dim Match As Boolean, LookFor As String
  FundNum$ = LTrim$(FundNum$)
  OpenAcctFile AcctFile
  NumOfAccts = LOF(AcctFile) / Len(GLAcct)
  For CntA = 1 To NumOfAccts
  Get AcctFile, CntA, GLAcct
    LookFor$ = Mid(GLAcct.Num, 1, Val(GLFundLen))
    If FundNum$ = LookFor$ Then
      If GLAcct.Deleted = 0 Then
        Match = True
        Close AcctFile
        Exit For
      End If
    End If
    Next
  If Match Then
    FindAcct = CntA
  Else
    FindAcct = 0
  End If
End Function
Public Function FindAcctFnct(FnctNum)
'To search for Account - see if Fund Has been used so can delete fund or not
  Dim NumOfAccts As Integer, CntA As Integer, AcctFile As Integer
  Dim Match As Boolean
  OpenAcctFile AcctFile
  NumOfAccts = LOF(AcctFile) / Len(GLAcct)
  For CntA = 1 To NumOfAccts
  Get AcctFile, CntA, GLAcct
    If FnctNum = GLAcct.FNCTRec Then
      If GLAcct.Deleted = 0 Then
        Match = True
        Close AcctFile
        Exit For
      End If
    End If
    Next
  If Match Then
    FindAcctFnct = CntA
  Else
    FindAcctFnct = 0
  End If
End Function

Public Function GetAcctNum$(RecordNumber)
  Dim AcctFileNum As Integer, NumAccts As Integer
   OpenAcctFile AcctFileNum, NumAccts
   If RecordNumber > 0 Then
     Get AcctFileNum, RecordNumber, GLAcct
     If GLAcct.Deleted = 0 Then
       GetAcctNum$ = GLAcct.Num
     Else
       GetAcctNum$ = "Invalid Acct"
     End If
   Else
     GetAcctNum$ = "Invalid Acct"
   End If
   Close AcctFileNum

End Function

Public Function GetAcctTitle$(RecordNumber)
  Dim AcctFileNum As Integer, NumAccts As Integer
   OpenAcctFile AcctFileNum, NumAccts
   If RecordNumber > 0 Then
     Get AcctFileNum, RecordNumber, GLAcct
     If GLAcct.Deleted = 0 Then
      GetAcctTitle$ = GLAcct.Title
     Else
      GetAcctTitle$ = "Invalid Acct"
     End If
   Else
     GetAcctTitle$ = "Invalid Acct"
   End If
   Close AcctFileNum

End Function

Public Function FindFund(FundNum$)
  Dim NumOfFunds As Integer, cnt As Integer, FundFile As Integer
  Dim Match As Boolean, LookFor As String
  FundNum$ = LTrim$(FundNum$)
  OpenFundFile FundFile, NumOfFunds
  'NumOfFunds = LOF(FundFile) / Len(GLFund)
  For cnt = 1 To NumOfFunds
  Get FundFile, cnt, GLFund
  LookFor$ = Trim$(GLFund.FundNum)
  If GLFund.Deleted = 0 Then
    If FundNum$ = LookFor$ Then
      Match = True
      Close FundFile
      Exit For
    End If
  End If
  Next
  If Match Then
    FindFund = cnt
  Else
    FindFund = 0
    Close FundFile
  End If
End Function
Public Function FindFnct(FnctNum$)
  Dim NumOfFncts As Integer, cnt As Integer, FnctFile As Integer
  Dim Match As Boolean, LookFor As String
  FnctNum$ = LTrim$(FnctNum$)
  OpenFnctFile FnctFile, NumOfFncts
  'NumOfFunds = LOF(FundFile) / Len(GLFund)
  For cnt = 1 To NumOfFncts
  Get FnctFile, cnt, GLFNCT
  LookFor$ = Trim$(GLFNCT.FnctNum)
  If GLFNCT.Deleted = 0 Then
    If FnctNum$ = LookFor$ Then
      Match = True
      Close FnctFile
      Exit For
    End If
  End If
  Next
  If Match Then
    FindFnct = cnt
  Else
    FindFnct = 0
    Close FnctFile
  End If
End Function

Public Function CheckValDate(ValCheck As String)
  Dim Month As Integer, Day As Integer, Year As Integer
  Month = Val(Mid(ValCheck, 1, 2))
  Day = Val(Mid(ValCheck, 4, 2))
  Year = Val(Mid(ValCheck, 7, 4))
  'Checks date if Blank then won't check for valid date
  'and then checks each section, month, day and year
  'if any section wrong then returns false value
      If InStr(ValCheck, "_") <= 0 Then
          If ((Month > 0) And (Month < 13)) Then
              If Day > 0 And Day < 32 Then
                  If Year > 1979 And Year < 2099 Then
                      CheckValDate = True
                  End If
              End If
          End If
      End If
End Function

Public Sub GetAcctStruct(GLUserName$, GLFundLen%, GLAcctLen%, GLDetLen%)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  GLUserName = QPTrim$(GLSetUpRec(1).UserName)
  GLFundLen = GLSetUpRec(1).FundLen
  GLAcctLen = GLSetUpRec(1).AcctLen
  GLDetLen = GLSetUpRec(1).DetLen
  Erase GLSetUpRec
End Sub
Public Function GetRPTName(Newrp As String)
  Dim Part As Double
  Part = Timer
  Newrp = Newrp + QPTrim(Str(CLng(Part))) + ".PRN"
End Function
Public Sub GetInvDef(POTabStop As Boolean, PSLDef As Boolean, DupInvDef As Boolean)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  POTabStop = GLSetUpRec(1).POStop
  If GLSetUpRec(1).DupInvFlag = 1 Then
    DupInvDef = True
  Else
    DupInvDef = False
  End If
  If GLSetUpRec(1).PSLFlag = 1 Then
    PSLDef = True
  Else
    PSLDef = False
  End If
  Erase GLSetUpRec
End Sub
Public Function SortFNCTIndex()
  Dim FnctIdxFileNum As Integer, NumFFIdxRecs As Integer, FnctFileNum As Integer
  Dim NumFncts As Integer, cnt As Long, GoodFncts As Long
  Dim OutOfOrder As Boolean, TempIdxRec As GLFNCTIndexType
  KillFile "GLFnct.IDX"
  OpenFnctIdx FnctIdxFileNum, NumFFIdxRecs
  OpenFnctFile FnctFileNum, NumFncts
  If NumFncts < 1 Then    'no need to sort if no record
    Close FnctIdxFileNum, FnctFileNum
    Exit Function
  End If
  ReDim Idxbuff(1 To NumFncts) As GLFNCTIndexType
  For cnt = 1 To NumFncts
    Get FnctFileNum, cnt, GLFNCT
    If GLFNCT.Deleted = 0 Then
      GoodFncts = GoodFncts + 1
      Idxbuff(GoodFncts).FnctNum = GLFNCT.FnctNum
      Idxbuff(GoodFncts).RecNum = cnt
    End If
  Next
  Close FnctFileNum
  If GoodFncts = 0 Then
    Close FnctIdxFileNum
    Exit Function
  End If
  ReDim Preserve Idxbuff(1 To GoodFncts) As GLFNCTIndexType
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To GoodFncts - 1
      If Idxbuff(cnt).FnctNum > Idxbuff(cnt + 1).FnctNum Then
        LSet TempIdxRec = Idxbuff(cnt)
        LSet Idxbuff(cnt) = Idxbuff(cnt + 1)
        LSet Idxbuff(cnt + 1) = TempIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
  For cnt = 1 To GoodFncts
    Put FnctIdxFileNum, cnt, Idxbuff(cnt)
  Next
  Close FnctIdxFileNum
End Function


Public Function SortFundIndex()
  Dim FundIdxFileNum As Integer, NumFFIdxRecs As Integer, FundFileNum As Integer
  Dim NumFunds As Integer, cnt As Integer, GoodFunds As Integer
  Dim OutOfOrder As Boolean, TempIdxRec As GLFundIndexType
  KillFile "GLFund.IDX"
  OpenFundIdx FundIdxFileNum, NumFFIdxRecs
  OpenFundFile FundFileNum, NumFunds
  If NumFunds < 1 Then    'no need to sort if no record
    Close FundIdxFileNum, FundFileNum
    Exit Function
  End If
  ReDim Idxbuff(1 To NumFunds) As GLFundIndexType
  For cnt = 1 To NumFunds
    Get FundFileNum, cnt, GLFund
    If GLFund.Deleted = 0 Then
      GoodFunds = GoodFunds + 1
      Idxbuff(GoodFunds).FundNum = GLFund.FundNum
      Idxbuff(GoodFunds).RecNum = cnt
    End If
  Next
  Close FundFileNum
  If GoodFunds = 0 Then
    Close FundIdxFileNum
    Exit Function
  End If
  ReDim Preserve Idxbuff(1 To GoodFunds) As GLFundIndexType
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To GoodFunds - 1
      If Idxbuff(cnt).FundNum > Idxbuff(cnt + 1).FundNum Then
        LSet TempIdxRec = Idxbuff(cnt)
        LSet Idxbuff(cnt) = Idxbuff(cnt + 1)
        LSet Idxbuff(cnt + 1) = TempIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
  For cnt = 1 To GoodFunds
    Put FundIdxFileNum, cnt, Idxbuff(cnt)
  Next
  Close FundIdxFileNum
End Function
Public Sub KillFileD(FileName$)
  On Local Error GoTo ErrorCatch
  If ExistD(FileName$) Then
    Kill FileName$
  End If
  Exit Sub
ErrorCatch:
  Select Case Err
    Case Is <> 53
      MainLog ("KillfileD error code is " + Str$(Err) + " .")
       MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
       GLTerminate
    Case 53
      Resume ExitFillFile
  End Select
    
ExitFillFile:
  
End Sub

Public Function ExistD(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  On Error GoTo LOGTHIS
  FileHandle = FreeFile
  Open FileName$ For Binary Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    ExistD = True
  Else
    ExistD = False
    Kill FileName$
    MainLog ("ExistD NOT-File " + FileName$ + "##@ Does not exist.")
  End If
  Exit Function
  
LOGTHIS:
  Call MainLog("ExistD Problem with " & FileName$)
  Resume Next
End Function
Public Function Exist(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  On Error GoTo LOGTHIS
  FileHandle = FreeFile
  Open FileName$ For Input Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    If FileName$ = "GLAcct.IDX" Then Kill FileName$
    Kill FileName$
    MainLog ("Exist NOT-File " + FileName$ + "##@ Does not exist.")
  End If
  Exit Function
  
LOGTHIS:
  Call MainLog("Problem with " & FileName$)
  Resume Next
End Function
Public Sub KillFile(FileName$)
Dim xxonce As Integer
xxonce = 0
  On Local Error GoTo ErrorCatch
tryagain:
  If Exist(FileName$) Then
    Kill FileName$
  End If
  Exit Sub
  'In wrightsville they were the error below when adding glaccts added retry and do not to terminate.
ErrorCatch:
  Select Case Err
    Case Is <> 53
      xxonce = xxonce + 1
      MainLog ("Killfile error code is " + Str$(Err) + " .")
      If FileName$ <> "GLAcct.IDX" Then
       MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
       GLTerminate
      Else
        If xxonce < 10 Then
          Resume tryagain
        Else
          MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
        End If
      End If
    Case 53
      Resume ExitFillFile
  End Select
    
ExitFillFile:
  
End Sub
Public Sub GLTerminate()
  Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  On Local Error Resume Next
  MainLog "GL Exited: "
  Ready4others PWcnt
  'If DebugMode = False Then
    Shell "CitiPak.exe", vbMaximizedFocus
  'End If
 ' DoTheTime
  DoEvents
  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
    DoEvents
    Unload Forms(UBFrmCnt)
  Next
  End
End Sub
Public Sub CitiTerminate()
  Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  ClearInUse PWcnt
  DoEvents
  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(UBFrmCnt)
  Next
  DoEvents
  End
End Sub

Public Function FileSize(FileName$)
  Dim FileHandle As Integer
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle

End Function
Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   Load frmLoadingRpt
   frmViewPrint.ReportName = ReportFile$
   frmViewPrint.Caption = Title
   frmViewPrint.PgNum = PgNum
   If ForceSBar Then
     frmViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmViewPrint.cmdAlignment.Enabled = True
     frmViewPrint.AlignRpt = AlgnRptfile$
    Else
      frmViewPrint.cmdAlignment.Enabled = False
    End If
   frmViewPrint.NoPbox = False
   Unload frmLoadingRpt
   frmViewPrint.Show 1
End Sub

Public Sub ViewPrnChks(ReportFile As String, DefPrinter As String, Optional ForceSBar As Boolean)
   frmLoadingRpt.Show
   frmViewPrint.ReportName = ReportFile$
   If ForceSBar Then
     frmViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   frmViewPrint.cmdAlignment.Enabled = False
   frmViewPrint.cmdAlignment.Visible = False
   frmViewPrint.cmdPrnScn.Enabled = False
   frmViewPrint.cmdPrnScn.Visible = False
   frmViewPrint.NoPbox = True
   frmViewPrint.thePrn = DefPrinter
   Unload frmLoadingRpt
   frmViewPrint.Show 1
End Sub

 Public Function AcctFind(AcctNum$)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim CntA As Integer, FirstRec As Integer, LastRec As Integer
  Dim Match As Boolean, LookFor As String, MiddleRec As Integer
  'AcctNum$ = QPTrim$(AcctNum$)
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
   Match = False
   FirstRec = 1
   LastRec = NumAIdxRecs
   LookFor$ = QPTrim$(AcctNum$)
    Do Until LastRec < FirstRec
      MiddleRec = (LastRec + FirstRec) \ 2
      
      Get AcctIdxFileNum, MiddleRec, GLAcctidx
      If LookFor$ = Trim$(GLAcctidx.AcctNum) Then
         CntA = GLAcctidx.RecNum
         Match = True
         Close AcctIdxFileNum
         Exit Do
      ElseIf LookFor$ < GLAcctidx.AcctNum Then
         LastRec = MiddleRec - 1
      Else
         FirstRec = MiddleRec + 1
      End If
   Loop

  If Match Then
    AcctFind = CntA
  Else
    Close AcctIdxFileNum
    AcctFind = 0
  End If
End Function
'Public Function SortAcctIndex(formname As Form)
'  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, AcctFileNum As Integer
'  Dim NumAccts As Integer, CntAc As Integer, GoodAccts As Integer
'  Dim OutOfOrder As Boolean, TempIdxRec As GLAcctIndexType
'  KillFile "GLAcct.IDX"
'  FrmShowPctComp.Label1 = "Initializing Account Index."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , formname
'  DoEvents
'  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'  OpenAcctFile AcctFileNum
'  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
'  If NumAccts <= 1 Then    'no need to sort one record
'    Close AcctIdxFileNum, AcctFileNum
'    Exit Function
'  End If
'  ReDim Idxbuff(1 To NumAccts) As GLAcctIndexType
'  For CntAc = 1 To NumAccts
'    FrmShowPctComp.ShowPctComp CntAc, NumAccts
'    Get AcctFileNum, CntAc, GLAcct
'    If GLAcct.Deleted = 0 Then
'      GoodAccts = GoodAccts + 1
'      Idxbuff(GoodAccts).AcctNum = GLAcct.Num
'      Idxbuff(GoodAccts).RecNum = CntAc
'    End If
'  Next
'  Close AcctFileNum
'  If GoodAccts = 0 Then
'    Close AcctIdxFileNum
'    Exit Function
'  End If
'  ReDim Preserve Idxbuff(1 To GoodAccts) As GLAcctIndexType
'  FrmShowPctComp.Label1 = "Sorting Accounts...Please Wait..."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , formname
'  DoEvents
'
'  FrmShowPctComp.ShowPctComp 15, 100
'  Do
'
'    OutOfOrder = False          'assume it's sorted
'    For CntAc = 1 To GoodAccts - 1
'      'FrmShowPctComp.ShowAcct IdxBuff(CntAc).AcctNum
'
'      If Idxbuff(CntAc).AcctNum > Idxbuff(CntAc + 1).AcctNum Then
'        LSet TempIdxRec = Idxbuff(CntAc)
'        LSet Idxbuff(CntAc) = Idxbuff(CntAc + 1)
'        LSet Idxbuff(CntAc + 1) = TempIdxRec
'        OutOfOrder = True       'we're not done yet
'      End If
'
'    Next
'  Loop While OutOfOrder
'
'  FrmShowPctComp.ShowPctComp 95, 100
'  For CntAc = 1 To GoodAccts
'    FrmShowPctComp.ShowPctComp CntAc, GoodAccts
'    Put AcctIdxFileNum, CntAc, Idxbuff(CntAc)
'  Next
'  Close AcctIdxFileNum
'End Function
Public Function QSortAcctIndex(formname As Form)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, AcctFileNum As Integer
  Dim NumAccts As Integer, CntAc As Integer, GoodAccts As Integer
  Dim OutOfOrder As Boolean, TempIdxRec As GLAcctIndexType
  Dim lngCurLow As Long, lngCurHigh As Long
  Dim i As Integer, j As Integer
  
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  If NumAccts < 1 Then    'no need to sort no records
    Close AcctFileNum
    Exit Function
  End If
  KillFile "GLAcct.IDX"
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  FrmShowPctComp.Label1 = "Initializing Account Index."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents

  ReDim Idxbuff(1 To NumAccts) As GLAcctIndexType
  For CntAc = 1 To NumAccts
    FrmShowPctComp.ShowPctComp CntAc, NumAccts
    Get AcctFileNum, CntAc, GLAcct
    If GLAcct.Deleted = 0 Then
      GoodAccts = GoodAccts + 1
      Idxbuff(GoodAccts).AcctNum = GLAcct.Num
      Idxbuff(GoodAccts).RecNum = CntAc
    End If
  Next
  Close AcctFileNum
  If GoodAccts = 0 Then
    Close AcctIdxFileNum
    Exit Function
  End If
  ReDim Preserve Idxbuff(1 To GoodAccts) As GLAcctIndexType
  FrmShowPctComp.Label1 = "Sorting Accounts...Please Wait..."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents
  lngCurLow = LBound(Idxbuff)
  lngCurHigh = UBound(Idxbuff)
  FrmShowPctComp.ShowPctComp 15, 100
  QSort Idxbuff(), lngCurLow, lngCurHigh
  
'    OutOfOrder = False          'assume it's sorted
'    For CntAc = 1 To GoodAccts - 1
'      'FrmShowPctComp.ShowAcct IdxBuff(CntAc).AcctNum
'
'      If Idxbuff(CntAc).AcctNum > Idxbuff(CntAc + 1).AcctNum Then
'        LSet TempIdxRec = Idxbuff(CntAc)
'        LSet Idxbuff(CntAc) = Idxbuff(CntAc + 1)
'        LSet Idxbuff(CntAc + 1) = TempIdxRec
'        OutOfOrder = True       'we're not done yet
'      End If
'
'    Next
'  Loop While OutOfOrder

  FrmShowPctComp.ShowPctComp 95, 100
  For CntAc = 1 To GoodAccts
    FrmShowPctComp.ShowPctComp CntAc, GoodAccts
    Put AcctIdxFileNum, CntAc, Idxbuff(CntAc)
  Next
  Close AcctIdxFileNum
End Function
Public Sub QSort(Idxbuff() As GLAcctIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As GLAcctIndexType
  Dim Temp2 As GLAcctIndexType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).AcctNum < Temp.AcctNum
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.AcctNum < Idxbuff(lngCurHigh).AcctNum
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub
Public Sub QCSort(Idxbuff() As ChkSortType, lLBound, lUBound)
'this is for check list sort by check or vendor
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As ChkSortType
  Dim Temp2 As ChkSortType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).CHKinfo < Temp.CHKinfo
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.CHKinfo < Idxbuff(lngCurHigh).CHKinfo
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QCSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QCSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub

Public Function FindDept(DeptNum$)
  Dim NumDIdxRecs As Integer, cnt As Integer, DeptIdxFileNum As Integer
  Dim Match As Boolean, LookFor As String
  DeptNum$ = LTrim$(DeptNum$)
  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
  For cnt = 1 To NumDIdxRecs
  Get DeptIdxFileNum, cnt, GLDeptIdx
  LookFor$ = Trim$(GLDeptIdx.DeptNum)
  If DeptNum$ = LookFor$ Then
     Match = True
     cnt = GLDeptIdx.RecNum
     Close DeptIdxFileNum
     Exit For
  End If
  Next
  If Match Then
    FindDept = cnt
  Else
    FindDept = 0
    Close DeptIdxFileNum
  End If
End Function
Public Function SortDeptIndex()
  Dim DeptIdxFileNum As Integer, NumDIdxRecs As Integer, DeptFileNum As Integer
  Dim NumDepts As Integer, cnt As Integer, GoodDepts As Integer
  Dim OutOfOrder As Boolean, TempIdxRec As GLDeptIndexType
  KillFile "GLDept.IDX"
  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
  OpenDeptFile DeptFileNum, NumDepts
  NumDepts = LOF(DeptFileNum) / Len(GLDept)
  If NumDepts < 1 Then    'no need to sort if no record
    Close DeptIdxFileNum, DeptFileNum
    Exit Function
  End If
  
  ReDim Idxbuff(1 To NumDepts) As GLDeptIndexType
  For cnt = 1 To NumDepts
    Get DeptFileNum, cnt, GLDept
    If GLDept.Deleted = 0 Then
      GoodDepts = GoodDepts + 1
      Idxbuff(GoodDepts).DeptNum = GLDept.DeptNum
      Idxbuff(GoodDepts).RecNum = cnt
    End If
  Next
  Close DeptFileNum
  
  If GoodDepts = 0 Then
    Close DeptIdxFileNum
    Exit Function
  End If
  
  ReDim Preserve Idxbuff(1 To GoodDepts) As GLDeptIndexType
  
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To GoodDepts - 1
      If Idxbuff(cnt).DeptNum > Idxbuff(cnt + 1).DeptNum Then
        LSet TempIdxRec = Idxbuff(cnt)
        LSet Idxbuff(cnt) = Idxbuff(cnt + 1)
        LSet Idxbuff(cnt + 1) = TempIdxRec
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
    
  For cnt = 1 To GoodDepts
    Put DeptIdxFileNum, cnt, Idxbuff(cnt)
  Next
  
  Close DeptIdxFileNum
End Function

Public Static Sub SmallPause()
Dim St1#, St2#
St1 = Timer
St2 = St1 + 0.0003
Do
  DoEvents
Loop Until Timer > St2

End Sub

Public Function InstrCount(ByVal strToCheck As String, ByVal strToFind As String) As Long
'On Error Resume Next

  Dim lngCount As Long
  Dim lngPos As Long
  
  lngCount = 0
  lngPos = 0
  Do
    lngPos = InStr(lngPos + 1, strToCheck, strToFind)
    If (lngPos > 0) Then
      lngCount = lngCount + 1
    End If
  Loop Until (lngPos = 0)
  
  InstrCount = lngCount
End Function
Public Function FillAcctList(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer
  Dim NumAIdxRecs As Integer
  Dim cnt As Integer
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  For cnt = 1 To NumAIdxRecs
    Get AcctIdxFileNum, cnt, GLAcctidx
    txtField.AddItem Trim(GLAcctidx.AcctNum)
  Next
  Close AcctIdxFileNum
End Function
Public Static Function FillFNCTList(txtField As fpCombo)
  Dim FnctIdxFileNum As Integer, NumFFIdxRecs As Integer, FnctFileNum As Integer
  Dim NumFncts As Integer, cnt As Long
  OpenFnctIdx FnctIdxFileNum, NumFFIdxRecs
  OpenFnctFile FnctFileNum, NumFncts
  NumFncts = LOF(FnctFileNum) / Len(GLFNCT)
  txtField.Row = -1
  For cnt = 1 To NumFFIdxRecs
    Get FnctIdxFileNum, cnt, GLFNCTIdx
    Get FnctFileNum, GLFNCTIdx.RecNum, GLFNCT
      If GLFNCT.Deleted = 0 Then
        txtField.InsertRow = QPTrim(GLFNCT.FnctNum) & Chr$(9) & Trim(GLFNCT.Title) & Chr$(9) & Str$(GLFNCTIdx.RecNum)
      End If
  Next
  Close FnctIdxFileNum
  Close FnctFileNum
  'Erase AcctIdxFileNum, NumAIdxRecs
  'Erase AcctFile, NumAccts, CntA
  End Function
Public Function GetFnctCode(RecordNum As Long)
  Dim FnctFile As Integer, NumFncts As Integer
  OpenFnctFile FnctFile, NumFncts
  Get FnctFile, RecordNum, GLFNCT
  GetFnctCode = QPTrim$(GLFNCT.FnctNum)
  'txtTitle = Trim(GLFNCT.Title)
  
  Close FnctFile
End Function
Public Function GetFnctTitle(RecordNum As Long)
  Dim FnctFile As Integer, NumFncts As Integer
  OpenFnctFile FnctFile, NumFncts
  Get FnctFile, RecordNum, GLFNCT
  'GetFnctCode = QPTrim$(GLFNCT.FnctNum)
  GetFnctTitle = QPTrim(GLFNCT.Title)
  
  Close FnctFile
End Function

Public Static Function FillAcctNumName(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField.Row = -1
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title) & Chr$(9) & QPStrip(GLAcct.Num)
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  'Erase AcctIdxFileNum, NumAIdxRecs
  'Erase AcctFile, NumAccts, CntA
  End Function
  Public Static Function FillAcctstwo(txtField1 As fpCombo, txtField2 As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  Dim TempList As String
  OpenAcctFile AcctFile, NumAccts
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  'NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField1.Row = -1
  txtField2.Row = -1
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        'Tried this for one column combo
        'TempList = (QPTrim(GLAcct.Num)) & "   " & Trim(GLAcct.Title)
        TempList = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title) & Chr$(9) & QPStrip(GLAcct.Num)
        'TempList = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title)
        txtField1.InsertRow = TempList
        txtField2.InsertRow = TempList
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  'Erase AcctIdxFileNum, NumAIdxRecs
  'Erase AcctFile, NumAccts, CntA
  End Function

  Public Function FundList(txtField As fpCombo)
  Dim FundIdxFileNum As Integer, NumFIdxRecs As Integer, cnt As Integer
  Dim FundFileNum As Integer, NumFunds As Integer
  OpenFundFile FundFileNum, NumFunds
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  txtField.InsertRow = Str$(0) & Chr(9) & ("ALL")
  For cnt = 1 To NumFIdxRecs
    Get FundIdxFileNum, cnt, GLFundIdx
    Get FundFileNum, GLFundIdx.RecNum, GLFund
      If GLFund.Deleted = 0 Then
           '''txtField.InsertRow = Str$(GLFund.FundNum) & Chr$(32) & QPTrim(GLFund.Title)

        txtField.InsertRow = (GLFund.FundNum) & Chr$(9) & QPTrim(GLFund.Title)
      End If
  Next
  Close FundIdxFileNum
  Close FundFileNum
End Function
Public Function DeptList(txtField As fpCombo)
  Dim DeptIdxFileNum As Integer, NumDIdxRecs As Integer, DeptFileNum As Integer
  Dim NumDepts As Integer, cnt As Integer
  OpenDeptFile DeptFileNum, NumDepts
  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
  txtField.InsertRow = Str$(0) & Chr$(9) & ("All") & Chr$(9) & ("Departments")
  For cnt = 1 To NumDIdxRecs
    Get DeptIdxFileNum, cnt, GLDeptIdx
    Get DeptFileNum, GLDeptIdx.RecNum, GLDept
      If GLDept.Deleted = 0 Then
        txtField.InsertRow = Str$(GLDeptIdx.RecNum) & Chr$(9) & QPTrim(GLDept.DeptNum) & Chr$(9) & QPTrim(GLDept.Title)
      End If
  Next
  Close DeptIdxFileNum
  Close DeptFileNum
End Function
Public Function GetFBAcct(FBAcct As String)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  FBAcct = QPTrim(GLSetup.FBAcct)
  Close SetupFile
End Function

  Public Function GetPostDates(LPDate As Integer, HPDate As Integer)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  LPDate = GLSetup.LPDate
  HPDate = GLSetup.HPDate
  Close SetupFile
End Function
Public Static Function Using$(ByVal fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(FmtLen)
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen = 0 Then
    TempNumber = "0"
    GoTo Gotazero
  End If
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
Gotazero:

  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function
Public Static Function Using0$(ByVal fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If Number > 0 Then
    If BuckPos = 1 Then
      fmt = Right$(fmt, FmtLen - 1)
    ElseIf BuckPos > 1 Then
      fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
    End If
    FmtNumber = Space$(FmtLen)
    TempNumber = Format(Number, fmt)
    TempLen = Len(TempNumber)
    If TempLen = 0 Then
      TempNumber = "0"
      GoTo Gotazero
    End If
    If TempLen >= 2 Then
      If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
        TempNumber = TempNumber + "0"
      End If
    End If
    If Right$(TempNumber, 1) = "." Then
      TempNumber = TempNumber + "00"
    End If
    If BuckPos > 0 Then
      TempNumber = "$" + TempNumber
    End If
  Else
    Using0 = ""
    Exit Function
  End If
Gotazero:

  RSet FmtNumber = TempNumber
  Using0 = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function

Public Function QPTrim$(Text As String)
  'Dim CPos As Long
  Dim StrLen As Long
  Dim cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function
Public Function QPStrip$(AcctNum$)
  Dim x As String, DashPos As Integer
   x$ = QPTrim$(AcctNum$)  '(Form$(AcctNum, 0))
   Do
      DashPos = InStr(x$, "-")
      If DashPos > 0 Then
         x$ = Left$(x$, DashPos - 1) + Mid$(x$, DashPos + 1)
      End If
    Loop While DashPos

    QPStrip$ = x$

End Function
'****************************************************************************
'formats an account number string with dashes.
'****************************************************************************
'
Public Function FmtAcct$(AN$, FundLen%, AcctLen%, DetLen%)
  Dim FmtTotAcctLen As Integer, ANLen As Integer
'   AN$ = QPTrim$(AN$)
'   FmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "

  FmtTotAcctLen = FundLen + AcctLen + DetLen

  AN$ = QPTrim$(AN$)
  ANLen = Len(AN$)

  If ANLen > FmtTotAcctLen Then
    AN$ = Left$(AN$, FmtTotAcctLen)
    ANLen = FmtTotAcctLen
  End If

  Select Case ANLen
    Case Is < FundLen
      FmtAcct$ = AN$
    Case FundLen
      FmtAcct$ = AN$ + "-"
    Case (FundLen + 1) To (AcctLen + FundLen) - 1
      FmtAcct$ = Left$(AN$, FundLen) + "-" + Mid$(AN$, FundLen + 1)
    Case (AcctLen + FundLen)
      FmtAcct$ = Left$(AN$, FundLen) + "-" + Mid$(AN$, FundLen + 1, AcctLen) + "-"
    Case (AcctLen + FundLen + 1) To (AcctLen + FundLen + DetLen) - 1
      FmtAcct$ = Left$(AN$, FundLen) + "-" + Mid$(AN$, FundLen + 1, AcctLen) + "-" + Mid$(AN$, FundLen + AcctLen + 1)
    Case (AcctLen + FundLen + DetLen)
      FmtAcct$ = Left$(AN$, FundLen) + "-" + Mid$(AN$, FundLen + 1, AcctLen) + "-" + Mid$(AN$, FundLen + AcctLen + 1, DetLen) 'RIGHT$(AN$, DetLen)
  End Select

End Function

Public Function Post2GL(FileName$, BadTrans, formname As Form, go4it As Boolean)
'****************************************************************************
' Input: FileName$ is the edit file to be posted, which is in the same type
'        as the transaction history (GLTRANS.DAT) file
' BadTrans returns the record number of a transaction which was not posted
'****************************************************************************
  On Local Error GoTo ItsBroke
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
  Dim NumAccts As Integer, AcctFileNum As Integer, Prev As Long, TransPosted As Long
  'SHARED Acct AS GLAcctRecType, Trans AS GLTransRecType
  Dim Tran2Post As GLTransRecType        'Dim a buffer for the edit file
  Dim TrRecLen As Integer, File2Post As Integer, Num2Post As Long
  Dim TransFileNum As Integer, NumTrans As Long, cnt As Long
  TrRecLen = Len(Tran2Post)              'Determine the rec length
  File2Post = FreeFile                   'Get a handle
  Open FileName$ For Random As File2Post Len = TrRecLen
  Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions
  Dim GLLogFileName As String, GLLogFile As Integer, Log As String
  Dim RecNum As Integer, DrPosted As Double, CrPosted As Double, Posted As Long
  Dim PRNFile As Integer, ReportFile As String, ToPrint As String
 'Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
 'Dim CntA As Integer, AcctNum As String, LookFor As String
  'OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
 ' ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
  'For CntA = 1 To NumAIdxRecs
 '   Get AcctIdxFileNum, CntA, IdxAry(CntA)
 ' Next
 '   Close AcctIdxFileNum
 
   If go4it = True Then
     FrmShowPctComp.Label1 = "Posting Account Transactions."
   Else
     FrmShowPctComp.Label1 = "Verifying Account Transactions."
   End If
   FrmShowPctComp.CmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
   DoEvents
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans
  '--update the posting log file
  If go4it = True Then
    GLLogFileName = "GLLog.dat"
    GLLogFile = FreeFile
    Open GLLogFileName$ For Append As GLLogFile
    Print #GLLogFile, "Post to General Ledger initiated on " + Date$ + " @ " + Time$
  Else
   PRNFile = FreeFile
   ReportFile$ = "TempLog.PRN"
   Open ReportFile$ For Output As #PRNFile
  End If
  Log = Space$(132)
  For cnt = 1 To Num2Post                'Start processing transactions
    FrmShowPctComp.ShowPctComp cnt, Num2Post
    Get File2Post, cnt, Tran2Post 'Get records from work file
     If Tran2Post.Marked = False Then
       RecNum = AcctFind(Tran2Post.AcctNum)   'Verify account is in G/L
       ' AcctNum$ = Trim$(Tran2Post.AcctNum)
'****Make The Find Faster!!!!
'-Find the record number of the account
'      For CntA = 1 To NumAIdxRecs
'        'Here you put Jump Around Code To Speed UP MOre!!!
'        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
'        If AcctNum$ = LookFor$ Then
'          RecNum = IdxAry(CntA).RecNum
'          Exit For
'        End If
'      Next

        If RecNum > 0 Then                  'if valid acct then proceed
           Get AcctFileNum, RecNum, GLAcct    'Get the account
           '--depending on account type, update running balance
           Select Case GLAcct.Typ
            Case "A", "E"                 'asset, exp accts
              GLAcct.Bal = ((GLAcct.Bal + Tran2Post.DrAmt) - (Tran2Post.CrAmt))
              If go4it = True Then
                Put AcctFileNum, RecNum, GLAcct
              End If
            Case "L", "R"                 'liab, rev accts
              GLAcct.Bal = ((GLAcct.Bal + Tran2Post.CrAmt) - (Tran2Post.DrAmt))
              If go4it = True Then
                Put AcctFileNum, RecNum, GLAcct
              End If
           End Select
           DrPosted = (DrPosted + Tran2Post.DrAmt)
           CrPosted = (CrPosted + Tran2Post.CrAmt)
           NumTrans = NumTrans + 1          'increment record pointer
           Get TransFileNum, NumTrans, GLTrans
           GLTrans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
           GLTrans.TRDATE = Tran2Post.TRDATE
           GLTrans.Desc = Tran2Post.Desc
           GLTrans.LDesc = Tran2Post.LDesc
           GLTrans.CrAmt = Tran2Post.CrAmt
           GLTrans.DrAmt = Tran2Post.DrAmt
           GLTrans.Ref = Tran2Post.Ref
           GLTrans.Src = Tran2Post.Src
           GLTrans.ChkByte = Chr$(1) 'this is for v205
           GLTrans.NextTran = 0
           If go4it = True Then
             Put TransFileNum, NumTrans, GLTrans
           End If
           Posted = Posted + 1
           Tran2Post.Marked = True
           If go4it = True Then
             Put File2Post, cnt, Tran2Post
           End If
           '---------------------------------Start linking here
           '--if first trans for this acct,
           If GLAcct.FrstTran = 0 Then
              GLAcct.FrstTran = NumTrans      'assign first & last pointers to
              GLAcct.LastTran = NumTrans      'this transaction
              If go4it = True Then
                Put AcctFileNum, RecNum, GLAcct
              End If
           '--Prior Transactions have been posted to this acct
           Else
                                                'in the account file..
              Prev = GLAcct.LastTran             'remember the prev trans point
              GLAcct.LastTran = NumTrans         'reset last trans to this tran
              If go4it = True Then
                Put AcctFileNum, RecNum, GLAcct
              End If                                  'In the trans file...
              Get TransFileNum, Prev, GLTrans    'Get the last transaction
              GLTrans.NextTran = NumTrans        'reset pointer to this trans
              If go4it = True Then
                Put TransFileNum, Prev, GLTrans
              End If
           End If
           TransPosted = TransPosted + 1
        Else                                'Account NOT found!
           BadTrans = BadTrans + 1          'Pass info back to caller
           If go4it = True Then
            GoSub LogGLPostErr
           Else
            GoSub LogTempErr
           End If
        End If
      End If  '--marked test
   Next
   If Num2Post < 1 Then
    FrmShowPctComp.ShowPctComp 1, 1
   End If
  If go4it = True Then
      If BadTrans = 0 Then
        Print #GLLogFile, ("No Posting Errors. Posted Transaction Count :" + Using$("####", TransPosted))
      End If
      Print #GLLogFile, ("Debits Posted :" + Using$("##,###,###.##", DrPosted))
      Print #GLLogFile, ("Credits Posted :" + Using$("##,###,###.##", CrPosted))
      Print #GLLogFile, String$(78, "-")
   Else
      If BadTrans = 0 Then
        Print #PRNFile, ("No Errors Found. Transaction Count :" + Using$("####", TransPosted))
      End If
  End If
  Close
Exit Function
'was printing register and deleteing edit file here.
'Now do this in module that called this sub
LogGLPostErr:
   Print #GLLogFile, "Unposted Transaction"
   Print #GLLogFile, "Record Number  :"; Str$(cnt)
   Print #GLLogFile, "Account Number :"; Tran2Post.AcctNum
   Print #GLLogFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #GLLogFile, "Description    :"; Tran2Post.Desc
   Print #GLLogFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #GLLogFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #GLLogFile, "***"
Return
LogTempErr:
   Print #PRNFile, "Unpostable Transaction"
   Print #PRNFile, "Record Number  :"; Str$(cnt)
   Print #PRNFile, "Account Number :"; Tran2Post.AcctNum
   Print #PRNFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #PRNFile, "Description    :"; Tran2Post.Desc
   Print #PRNFile, "               :"; Tran2Post.LDesc
   Print #PRNFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #PRNFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #PRNFile, "***"
Return
ItsBroke:
  BadTrans = BadTrans + 1
  Print #PRNFile, "Error *** Call Software Support***"
  Print #PRNFile, "Record Number :"; Str$(cnt); Tran2Post.AcctNum
  Print #PRNFile, "Error Code"; Str(Err.Number)
  Resume Next
  
End Function
Public Function GetBankList(txtName As fpCombo)
  Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
  Dim Bank As GLBankRecType, BankRecLen As Integer, BankFile As Integer
  Dim BankNum As String, cnt As Integer, NumBanks As Integer
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  BankRecLen = Len(Bank)
  OpenBankFile BankFile, NumBanks
  If CDActive = "Y" Then
    Close BankFile
    txtName.InsertRow = (99 & Chr$(9) & "System Bank")
    Exit Function
  End If
  txtName.Row = -1
  For cnt = 1 To NumBanks
    Get BankFile, cnt, Bank
    If Bank.Deleted = 0 Then
      txtName.InsertRow = (Bank.BankNum & Chr$(9) & Mid$(Bank.BankName, 1, 25))
      '(Bank.BankNum & "  " & )
      'txtName. (txtName.NewIndex) = (Val(Bank.BankNum))
    Else
    End If
  Next
  Close BankFile
End Function

Public Sub FillallBanks(txtName As fpCombo)
  Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
  Dim Bank As GLBankRecType, BankRecLen As Integer, BankFile As Integer
  Dim BankNum As String, cnt As Integer, NumBanks As Integer
  BankRecLen = Len(Bank)
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  OpenBankFile BankFile, NumBanks
  txtName.Row = -1
  If CDActive = "Y" Then
    txtName.InsertRow = (99 & Chr$(9) & "System Bank")
  End If
  If NumBanks > 0 Then
  For cnt = 1 To NumBanks
    Get BankFile, cnt, Bank
    If Bank.Deleted = 0 Then
      txtName.InsertRow = (Bank.BankNum & Chr$(9) & Mid$(Bank.BankName, 1, 25))
    Else
    End If
  Next
  End If
  Close BankFile
End Sub
Public Function GetBankGLAcct(BankNum)
  Dim Bank As GLBankRecType, BankRecLen As Integer, BankFile As Integer
  Dim cnt As Integer, NumBanks As Integer
  Dim GLBank(1) As GLBankRecType
  BankRecLen = Len(GLBank(1))
  OpenBankFile BankFile, NumBanks
  Get BankFile, BankNum, GLBank(1)
  Close BankFile
  GetBankGLAcct = (GLBank(1).GLAcct)
End Function

Public Sub GetCentDep(CDActive$, CashAcct$, CDCash$, CDDue$)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  CDActive = GLSetUpRec(1).CDActive
  CashAcct = QPTrim(GLSetUpRec(1).CashAcct)
  CDCash = QPTrim(GLSetUpRec(1).CDCash)
  CDDue = QPTrim(GLSetUpRec(1).CDDue)
  
  Erase GLSetUpRec
End Sub
Public Sub SetDefBank(x$, DefBnk As Integer)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  ReDim GLSetUpRec(1) As GLSetupRecType
  SetUpRecLen = Len(GLSetUpRec(1))
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetUpRec(1)
  Close SetupFile
  Select Case x$
  Case "D"
    DefBnk = GLSetUpRec(1).CDBank
  Case "R"
    DefBnk = GLSetUpRec(1).CRBank
  Case "C"
    DefBnk = GLSetUpRec(1).ChkBank
  Case Else
  End Select

  
 Erase GLSetUpRec
End Sub

'Public Function MakeCRDate$(DateNum As Integer)
'Dim TempDate As String, SlashPos As Integer
'TempDate = Format(DateAdd("d", (DateNum), "12-31-1979"), "mm/dd/yyyy")
'Do
'  SlashPos = InStr(TempDate, "/")
'  If SlashPos > 0 Then
'    TempDate = Left$(TempDate, SlashPos - 1) + Mid$(TempDate, SlashPos + 1)
'  End If
'Loop While SlashPos > 0
'TempDate = Left$(TempDate, 6)
'MakeCRDate$ = TempDate
'End Function
'

 Public Function PostCJTrans(CJType As Integer, formname As Form)
  Dim CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, TotDr As Long, TotCr As Long, Active As Integer, BadTrans As Integer
  Dim PRNFile As Integer, RecLen As Integer, RecordNum As Integer, CntD As Integer
  Dim GLLogFileName As String, GLLogFile As Integer, Log As String, FundDue As String
  Dim ReportFile As String, ToPrint As String, DetPad As String, CshAcct As String
  Dim FundCode As String, OutofBal As Integer, PadChars As Integer, A As String, B As String
  Dim CommaFmt As String, FundNum As String, strMsg As String, MSrc As String
  Dim BadCashAcct As Boolean, PRNfileName As String, JEDebits As Double, JECredits As Double
  Dim Linecnt As Integer, IFFile As Integer, NumTrans As Integer
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  Dim Trans2Post As GLTransRecType
  Dim TmpSortTrans As GLTransRecType
  Dim OutOfOrder As Boolean, Editing As Boolean, TempCash As String
  Dim CJEdit As CJEditRecType
  Dim Tr2Post As GLTransRecType
  Dim TmpIFFile As String, CJPrnFile As String, CJEditFile As String
  Dim OSChek As OSChekRecType
  Dim OSChekFile As Integer, NumOSChks As Integer, OSChekFileNum As Integer, chk As Integer
  Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
'the get list of funds on gj main form
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  ReDim TrFundSum(1 To NumFunds) As Double
  CommaFmt$ = "##,###,###.##"
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Editing = False
'Set CJType for Correct file names 1 is Receipt, 2 is Disbursement
  Select Case CJType
    Case 1:
      CJEditFile$ = "GLCREd.dat"
      TmpIFFile$ = "CJRPOST.dat"
      MSrc$ = "CR" + Format$(Now, "mmddyy")
      CJPrnFile$ = "CRTrans.Prn"
    Case 2:
      CJEditFile$ = "GLCDEd.dat"
      TmpIFFile$ = "CJDPOST.dat"
      MSrc$ = "CD" + Format$(Now, "mmddyy")
      CJPrnFile$ = "CDTrans.Prn"
  End Select
 'Make sure have entries to post
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  For cnt = 1 To NumEdTrans
    Get CJEditFileNum, cnt, CJEdit
    If Not CJEdit.DelFlag Then
      Active = Active + 1
    End If
    If CJEdit.LOCKED = True Then
      Editing = True
      Exit For
    End If
  Next
  Close CJEditFileNum
'Give options to cancel posting
  If Not Editing Then
    If Active = 0 Then
      MsgBox "No Transactions To Post", vbOKOnly, "Post Canceled"
      GoSub OutBeforePost
    End If
    SetAttr (CJEditFile$), vbReadOnly
  Else
    MsgBox "The Cash Journal is being Edited, Please Close Edit Procedures Before trying to Post.", vbOKOnly, "Post Canceled"
    GoSub OutBeforePost
  End If
  If MsgBox("Are You Sure You Wish to Post.", vbOKCancel, "CD Posting") = vbCancel Then
    GoSub OutBeforePost
  Else
  'If Central Depository used then will need detail for acct #
    If CDActive$ = "Y" Then
      PadChars = GLDetLen - GLFundLen
      If PadChars > 0 Then
        DetPad$ = String(PadChars, "0")
      End If
    End If
    
    Active = 0                             'Reset Active counter for posting
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    IFFile = FreeFile
    Open TmpIFFile$ For Random As IFFile Len = Len(Tr2Post)
    RecordNum = LOF(IFFile) \ Len(Tr2Post)
    If RecordNum > 0 Then
      MsgBox "Interface File Already Exists. Possible Problems with Previous Posting." & Chr(13) & "Notify Software Support, DO NOT TRY POSTING AGAIN.", vbOKOnly, "Warning!!!"
      Close
      GoSub OutBeforePost
    End If
    RecordNum = 0
    For cnt = 1 To NumEdTrans              'Assign edit file to trans format
      For Fund = 1 To NumFunds
        TrFundSum#(Fund) = 0
      Next
      Get CJEditFileNum, cnt, CJEdit
      If Not CJEdit.DelFlag Then
        'If Central Depository use Central Deposi Cash Acct
       If CDActive$ = "Y" Then
        CshAcct$ = CashAcct$
        RecordNum = RecordNum + 1
        Tr2Post.AcctNum = CDCash$
        Tr2Post.TRDATE = CJEdit.TRDATE
        Tr2Post.Desc = CJEdit.Desc
        Tr2Post.LDesc = CJEdit.LDesc
        Tr2Post.Ref = CJEdit.DOCREF
        If CJType = 1 Then
          Tr2Post.DrAmt = Round(CJEdit.Amt)
          Tr2Post.CrAmt = 0
        ElseIf CJType = 2 Then
          Tr2Post.DrAmt = 0
          Tr2Post.CrAmt = Round(CJEdit.Amt)
        End If
        Tr2Post.Src = MSrc$
        Put #2, RecordNum, Tr2Post
      Else
        'Find Cash Acct Num from Bank Code
        If Exist("glbank.dat") Then
          TempCash$ = GetBankGLAcct(Val(CJEdit.RECCODE))
        End If
        If TempCash$ <> "" Then
          CshAcct$ = TempCash$
        Else
          CshAcct$ = CashAcct$
        End If
      End If
      'Add each Distribution to the interface file
      For CntD = 1 To 36
        If CJEdit.Dist(CntD).DACREC > 0 Then
          RecordNum = RecordNum + 1
          Tr2Post.AcctNum = CJEdit.Dist(CntD).DACN
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.Desc = CJEdit.Desc
          Tr2Post.LDesc = CJEdit.LDesc
          Tr2Post.Ref = CJEdit.DOCREF
          If CJType = 1 Then
            Tr2Post.DrAmt = 0
            Tr2Post.CrAmt = CJEdit.Dist(CntD).DAMT
          ElseIf CJType = 2 Then
            Tr2Post.DrAmt = CJEdit.Dist(CntD).DAMT
            Tr2Post.CrAmt = 0
          End If
          Tr2Post.Src = MSrc$
          Put #2, RecordNum, Tr2Post
          ' Add to Fund tot
          For Fund = 1 To NumFunds '- 1
            FundNum$ = Left$(CJEdit.Dist(CntD).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              TrFundSum#(Fund) = Round#(TrFundSum#(Fund) + CJEdit.Dist(CntD).DAMT)
              Exit For
            End If
          Next 'Fund
        End If
      Next 'Distribution
    'No More Distributions so now create the "Opposite" entries.
    'One to cash or central dep for each fund
      For Fund = 1 To NumFunds '- 1
        If TrFundSum#(Fund) <> 0 Then
        'if using Cent Dep Create Detail for Due to acct.
          If CDActive$ = "Y" Then
            If PadChars > 0 Then
              FundDue$ = FundList$(Fund) + DetPad$
            Else
              FundDue$ = FundList$(Fund)
            End If
          End If
        
        A$ = FundList$(Fund) + CshAcct$
        B = A$
        If AcctFind(A$) = 0 Then
          BadCashAcct = True
        End If
      'Fund's Cash or Cent Dep entry
        RecordNum = RecordNum + 1
        Tr2Post.AcctNum = A
        Tr2Post.TRDATE = CJEdit.TRDATE
        Tr2Post.Desc = CJEdit.Desc
        Tr2Post.LDesc = CJEdit.LDesc
        Tr2Post.Ref = CJEdit.DOCREF
        If CJType = 1 Then
          Tr2Post.DrAmt = TrFundSum#(Fund)
          Tr2Post.CrAmt = 0
        ElseIf CJType = 2 Then
          Tr2Post.DrAmt = 0
          Tr2Post.CrAmt = TrFundSum#(Fund)
        End If
        Tr2Post.Src = MSrc$
        Put #2, RecordNum, Tr2Post
    'Entry to Cent Dep Due to acct
          If CDActive$ = "Y" Then
            RecordNum = RecordNum + 1
            Tr2Post.AcctNum = (QPTrim(CDDue$) + FundDue$)
            Tr2Post.TRDATE = CJEdit.TRDATE
            Tr2Post.Desc = CJEdit.Desc
            Tr2Post.LDesc = CJEdit.LDesc
            Tr2Post.Ref = CJEdit.DOCREF
            If CJType = 1 Then
              Tr2Post.DrAmt = 0
              Tr2Post.CrAmt = TrFundSum#(Fund)
            ElseIf CJType = 2 Then
              Tr2Post.DrAmt = TrFundSum#(Fund)
              Tr2Post.CrAmt = 0
            End If
            Tr2Post.Src = MSrc$
            Put #2, RecordNum, Tr2Post
          End If
        End If
      Next
      End If
    Next
    If BadCashAcct Then
      Close
      SetAttr (CJEditFile$), vbNormal
      KillFile TmpIFFile$
      MsgBox "Invalid Cash Account, Posting Aborted. Check Journal Report for Invalid Entry.", vbOKOnly, "Posting Aborted"
      Exit Function
    End If
    Close
    Call Post2GL(TmpIFFile$, BadTrans, formname, False) 'common post & link sub
    If BadTrans <> 0 Then
      Close
      SetAttr (CJEditFile$), vbNormal
      KillFile TmpIFFile$
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Error Log"
      End If
      frmCitiCancel.Show
      Unload formname
      'Need to unload menu but how to reference it?????
      Exit Function
    End If
    If CJType = 2 Then
    'Post disbursements to o/s check file
      OpenOSChekFile OSChekFileNum, NumOSChks
      OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
      For chk = 1 To NumEdTrans
        Get CJEditFileNum, chk, CJEdit
        If Not CJEdit.DelFlag Then
          NumOSChks = NumOSChks + 1
          OSChek.ChkNum = Val(CJEdit.DOCREF)
          OSChek.chkdate = CJEdit.TRDATE
          OSChek.Bankcode = CJEdit.RECCODE
          OSChek.Desc = CJEdit.Desc
          OSChek.Amt = CJEdit.Amt
          OSChek.Src = 0 ' Why this ????? code as apcheck
          OSChek.Cleared = 0
          OSChek.VoidFlag = 0
          Put OSChekFileNum, NumOSChks, OSChek
        End If
      Next
    End If
    Close
    Call Post2GL(TmpIFFile$, BadTrans, formname, True) 'common post & link sub
    If BadTrans <> 0 Then                  'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      GLLogFileName = "GLlog.dat"
      ReportFile$ = "GLlog.dat"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Posting Log"
      End If
    End If
    frmReportOpt.Show 1
    If rptopt = 1 Then
      GoSub PrnPostJournal
    ElseIf rptopt = 2 Then
      GoSub PrnPostJournal2
    End If
    SetAttr (CJEditFile$), vbNormal
    KillFile CJEditFile$                    'kill the temp files
    KillFile TmpIFFile$

  MsgBox "Posting Procedure Completed", vbOKOnly, "Cash Journal Posting"
  End If
Exit Function
PrnPostJournal:
  RecLen = Len(Trans2Post)
  IFFile = FreeFile
  Open TmpIFFile$ For Random As IFFile Len = RecLen
  NumTrans = LOF(IFFile) \ RecLen
  ReDim SorTtrans(1 To NumTrans) As GLTransRecType
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    SorTtrans(cnt) = Trans2Post
  Next
  Close
  '*** What is SortT ??
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To NumTrans - 1
      If SorTtrans(cnt).AcctNum > SorTtrans(cnt + 1).AcctNum Then
        LSet TmpSortTrans = SorTtrans(cnt)
        LSet SorTtrans(cnt) = SorTtrans(cnt + 1)
        LSet SorTtrans(cnt + 1) = TmpSortTrans
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
  'The SortT below was from old Program **The Section Above (Per Dale) Replaced it.
  'SortT SortTrans(1), NumTrans, 0, 96, 2,14
  IFFile = FreeFile
  Open TmpIFFile$ For Random As IFFile Len = RecLen
  
  For cnt = 1 To NumTrans
    Put IFFile, cnt, SorTtrans(cnt)
  Next
  ToPrint$ = ""
  PRNFile = FreeFile
  PRNfileName$ = "CJPOST.PRN"
  Open PRNfileName$ For Output As #PRNFile
  
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    JEDebits# = Round#(JEDebits# + Trans2Post.DrAmt)
    JECredits# = Round#(JECredits# + Trans2Post.CrAmt)
   
    ToPrint$ = Format(DateAdd("d", (Trans2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    ToPrint$ = ToPrint$ + "~" + Trans2Post.AcctNum
    ToPrint$ = ToPrint$ + "~" + Left$(Trans2Post.Desc, 15) + " " + QPTrim$(Trans2Post.LDesc)
    ToPrint$ = ToPrint$ + "~" + Trans2Post.Ref
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Trans2Post.DrAmt)
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Trans2Post.CrAmt)
    Print #PRNFile, ToPrint$
  Next
  
  Close
  Load frmLoadingRpt
  If CJType = 1 Then
    ARptCashPost.Title = "Cash Receipt Journal Post Report"
  ElseIf CJType = 2 Then
    ARptCashPost.Title = "Cash Disbursement Journal Post Report"
  End If
  ARptCashPost.totDebit = Using$(CommaFmt$, JEDebits#)
  ARptCashPost.totCredit = Using$(CommaFmt$, JECredits#)
  ARptCashPost.txtDate = Now
  ARptCashPost.txtTown = GLUserName$
  ARptCashPost.GetName PRNfileName$
  ARptCashPost.startrpt

Return

PrnPostJournal2:
  RecLen = Len(Trans2Post)
  IFFile = FreeFile
  Open TmpIFFile$ For Random As IFFile Len = RecLen
  NumTrans = LOF(IFFile) \ RecLen
  ReDim SorTtrans(1 To NumTrans) As GLTransRecType
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    SorTtrans(cnt) = Trans2Post
  Next
  Close
  '*** What is SortT ??
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To NumTrans - 1
      If SorTtrans(cnt).AcctNum > SorTtrans(cnt + 1).AcctNum Then
        LSet TmpSortTrans = SorTtrans(cnt)
        LSet SorTtrans(cnt) = SorTtrans(cnt + 1)
        LSet SorTtrans(cnt + 1) = TmpSortTrans
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
  'The SortT below was from old Program **The Section Above (Per Dale) Replaced it.
  'SortT SortTrans(1), NumTrans, 0, 96, 2,14
  IFFile = FreeFile
  Open TmpIFFile$ For Random As IFFile Len = RecLen
  
  For cnt = 1 To NumTrans
    Put IFFile, cnt, SorTtrans(cnt)
  Next
  ToPrint$ = Space$(82)
  PRNFile = FreeFile
  PRNfileName$ = "CJPOST.PRN"
  Open PRNfileName$ For Output As #PRNFile
  GoSub CDJEHeader
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    JEDebits# = Round#(JEDebits# + Trans2Post.DrAmt)
    JECredits# = Round#(JECredits# + Trans2Post.CrAmt)
   
    LSet ToPrint$ = ""
    LSet ToPrint$ = Format(DateAdd("d", (Trans2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Mid$(ToPrint$, 13) = Trans2Post.AcctNum
    Mid$(ToPrint$, 29) = Left$(Trans2Post.Desc, 15)
    Mid$(ToPrint$, 42) = Trans2Post.Ref
    Mid$(ToPrint$, 50) = Using$(CommaFmt$, Trans2Post.DrAmt)
    Mid$(ToPrint$, 68) = Using$(CommaFmt$, Trans2Post.CrAmt)
    Print #PRNFile, ToPrint$
    Linecnt = Linecnt + 1
    If Len(QPTrim$(Trans2Post.LDesc)) > 0 Then
      Print #PRNFile, Tab(29); QPTrim$(Trans2Post.LDesc)
      Linecnt = Linecnt + 1
    End If
    If Linecnt > 55 Then
      Print #PRNFile, Chr$(12)
      GoSub CDJEHeader
    End If
  Next
  
  Print #PRNFile, ' Blank line
  Print #PRNFile, "Posting Totals:";
  Print #PRNFile, Tab(50); Using$(CommaFmt$, JEDebits#);
  Print #PRNFile, Tab(68); Using$(CommaFmt$, JECredits#);
  Print #PRNFile, Chr$(12)
  Close
  'File Must be closed before going to ViewPrint - Hence the close above
  If CJType = 1 Then
    ViewPrint PRNfileName$, "Cash Receipt Journal Post Report"
  ElseIf CJType = 2 Then
    ViewPrint PRNfileName$, "Cash Disbursement Journal Post Report"
  End If
Return
CDJEHeader:
  If CJType = 1 Then
    Print #PRNFile, "Cash Receipt Journal Entries"
  ElseIf CJType = 2 Then
    Print #PRNFile, "Cash Disbursement Journal Entries"
  End If
  Print #PRNFile, "Module: " + MSrc$
  Print #PRNFile,
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 1) = "Date"
  Mid$(ToPrint$, 15) = "Acct No"
  Mid$(ToPrint$, 29) = "Description"
  Mid$(ToPrint$, 41) = "Ref"
  Mid$(ToPrint$, 54) = "     Debit"
  Mid$(ToPrint$, 71) = "    Credit"
  Print #PRNFile, ToPrint$
  Print #PRNFile, String$(82, "=")
  Linecnt = 5
Return

OutBeforePost:
  SetAttr (CJEditFile$), vbNormal
  Exit Function
End Function


'****************************************************************************
'Retrieves the GL account type from the account data file.
'****************************************************************************
'
Public Function GetAcctType$(AcctRecNum)
  Dim AcctFileNum As Integer, NumAccts As Integer
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  Get AcctFileNum, AcctRecNum, GLAcct
  GetAcctType$ = GLAcct.Typ
  Close AcctFileNum
End Function

Public Sub GetFYDates(FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)
  Dim SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  FY1BegDate = GLSetup.FYBeg
  FY1EndDate = GLSetup.FYEnd
  FY2BegDate = GLSetup.NYBeg
  FY2EndDate = GLSetup.NYEnd
  Close SetupFile
End Sub

'***************************************************************************
'Finds the next undeleted record.
'Call with NextRec value of -1 for previous record, +1 for the next record.
'If a record is not found, the function returns the value of CurrRec.
'***************************************************************************
'
Function GetNextRec(FileNum, NumRecs, CurrRec, NextRec)
   Dim Found As Integer, Rec As Integer
   Found = 0
   Rec = CurrRec
   Do
      Rec = Rec + NextRec                'Set file pointer to next record
      If Rec > NumRecs Or Rec <= 0 Then  'test for beg or end of file
         Found = 0                       'if no more records then get out
         Exit Do
      End If
      Get FileNum, Rec, BgtEdit           'Get the record
      If BgtEdit.Deleted = 0 Then         'Ok if not deleted
         Found = 1
         Exit Do                         'Get out of loop when we find one
      End If
   Loop
   If Found = 0 Then
      GetNextRec = CurrRec
   Else
      GetNextRec = Rec
   End If
End Function
Public Sub GetFundList(FundList$(), NumFunds)
  Dim FundIndex As GLFundIndexType
  Dim FundIdxFile As Integer, cnt As Integer
  OpenFundIdx FundIdxFile, NumFunds
  If NumFunds = 0 Then
    MsgBox "No Funds", vbOKOnly, "No Funds"
    Close
    Exit Sub
  End If
  ReDim FundList$(1 To NumFunds)
  For cnt = 1 To NumFunds
    Get FundIdxFile, cnt, FundIndex
    FundList$(cnt) = Trim$(FundIndex.FundNum)
  Next
  Close FundIdxFile
End Sub
Public Sub GetFnctList(FnctList$(), NumFncts)
  Dim FnctIndex As GLFNCTIndexType
  Dim FnctIdxFile As Integer, cnt As Long
  OpenFnctIdx FnctIdxFile, NumFncts
  If NumFncts = 0 Then
    MsgBox "No Functions", vbOKOnly, "No Functions"
    Close
    Exit Sub
  End If
  ReDim FnctList$(1 To NumFncts)
  For cnt = 1 To NumFncts
    Get FnctIdxFile, cnt, FnctIndex
    FnctList$(cnt) = Trim$(FnctIndex.FnctNum)
  Next
  Close FnctIdxFile
End Sub

'****************************************************************************
'Retrieves the fund title from the fund data file.
'****************************************************************************
'
Public Function GetFundTitle(FundRecNum)
  Dim NumFunds As Integer, FundFileNum As Integer
  Dim GLFund As GLFundRecType
   OpenFundFile FundFileNum, NumFunds
   Get FundFileNum, FundRecNum, GLFund
   GetFundTitle = GLFund.Title
   Close FundFileNum

End Function

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function
Public Function RoundDol#(ByVal N#)
  RoundDol# = (Int(N# * 1 + 0.51)) / 1
End Function
Public Function BudAcctNumName(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField.Row = -1
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
          txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & Trim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title) & Chr$(9) & QPStrip(GLAcct.Num)
        End If
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  End Function
Public Function BudAcctstwo(txtField1 As fpCombo, txtField2 As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  Dim TempBudList As String
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField1.Row = -1
  txtField2.Row = -1
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, GLAcctidx
    Get AcctFile, GLAcctidx.RecNum, GLAcct
      If GLAcct.Deleted = 0 Then
        If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
          TempBudList = Str$(GLAcctidx.RecNum) & Chr$(9) & Trim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.Title) & Chr$(9) & QPStrip(GLAcct.Num)
          txtField1.InsertRow = TempBudList
          txtField2.InsertRow = TempBudList
        End If
      End If
  Next
  Close AcctIdxFileNum
  Close AcctFile
  End Function

Public Function SortT(Trsort() As TrSortType, NumAcctTrans)
  Dim TmpSort As TrSortType
  'ReDim Trsort(1 To 20000) As TrSortType
  Dim OutOfOrder As Boolean, cntT As Integer
      Do
        OutOfOrder = False          'assume it's sorted
        For cntT = 1 To NumAcctTrans - 1
          If Trsort(cntT).TRDATE > Trsort(cntT + 1).TRDATE Then
            LSet TmpSort = Trsort(cntT)
            LSet Trsort(cntT) = Trsort(cntT + 1)
            LSet Trsort(cntT + 1) = TmpSort
            OutOfOrder = True       'we're not done yet
          End If
        Next
      Loop While OutOfOrder

End Function
Public Function SortTRec(TranInfo() As TranRecInfoType, NumTrans)
  Dim TmpSort As TranRecInfoType
  Dim OutOfOrder As Boolean, cntT As Long
      Do
        OutOfOrder = False          'assume it's sorted
        For cntT = 1 To NumTrans - 1
          If TranInfo(cntT).TranDate > TranInfo(cntT + 1).TranDate Then
            LSet TmpSort = TranInfo(cntT)
            LSet TranInfo(cntT) = TranInfo(cntT + 1)
            LSet TranInfo(cntT + 1) = TmpSort
            OutOfOrder = True       'we're not done yet
          End If
        Next
      Loop While OutOfOrder

End Function

Public Function Num2Month%(Dt%)
  Dim d As String, M As String
  d$ = Format(DateAdd("d", Dt%, "12-31-1979"), "mm/dd/yyyy")
  M$ = Right$(d$, 2) + Left$(d$, 2)
  Num2Month% = Val(M$)

End Function

Function InQtr(TDate, RDate)
  Dim r As String, RM As String, T As String, TM As String
  Dim RQ As Integer, TQ As Integer
  '--Get the Report Quarter
  r$ = Format(DateAdd("d", RDate, "12-31-1979"), "mm/dd/yyyy")
  RM$ = Left$(r$, 2)

  '--Get the Transaction Quarter
  T$ = Format(DateAdd("d", TDate, "12-31-1979"), "mm/dd/yyyy")
  TM$ = Left$(T$, 2)

  Select Case RM$
    Case "01", "02", "03"
      RQ = 1
    Case "04", "05", "06"
      RQ = 2
    Case "07", "08", "09"
      RQ = 3
    Case "10", "11", "12"
      RQ = 4
  End Select

  Select Case TM$
    Case "01", "02", "03"
      TQ = 1
    Case "04", "05", "06"
      TQ = 2
    Case "07", "08", "09"
      TQ = 3
    Case "10", "11", "12"
      TQ = 4
  End Select

  If TQ = RQ Then
    InQtr = True
  Else
    InQtr = False
  End If

End Function

Public Function GetDeptTitle$(DeptRecNum)
  Dim DeptRec As GLDeptRecType
  Dim DeptFileNum As Integer, NumDepts As Integer
  OpenDeptFile DeptFileNum, NumDepts
  Get DeptFileNum, DeptRecNum, DeptRec
  GetDeptTitle$ = DeptRec.Title
  Close DeptFileNum

End Function

Public Function GetPct$(N1#, N2#)
  Dim Pct As String, P As String, PP As Double
  Pct$ = Space$(5)
  If N1# > 0 And N2# > 0 Then
    PP# = Round#((N1# / N2#) * 100)
    P$ = Str$(Int(PP#)) + "%"
    RSet Pct$ = P$
  End If
  GetPct$ = QPTrim(Pct$)

End Function
  Public Function FundstoList(x As fpCombo)
  Dim FundIdxFileNum As Integer, NumFIdxRecs As Integer, cnt As Integer
  Dim FundFileNum As Integer, NumFunds As Integer
  OpenFundFile FundFileNum, NumFunds
  OpenFundIdx FundIdxFileNum, NumFIdxRecs

  For cnt = 1 To NumFIdxRecs
    Get FundIdxFileNum, cnt, GLFundIdx
    Get FundFileNum, GLFundIdx.RecNum, GLFund
      If GLFund.Deleted = 0 Then
        x.AddItem (QPTrim(GLFund.FundNum)) & Chr$(9) & QPTrim(GLFund.Title) & Chr$(9) & QPTrim(GLFund.FundNum)
      End If
  Next
  Close FundIdxFileNum
  Close FundFileNum
End Function
Public Sub ReLinkPOTrans(x As Form, Optional NM As Boolean)
  Dim POTrans As GLTransRecType
  Dim Acct As GLAcctRecType
  Dim GLAcctFile As Integer, NumAccts As Integer
  Dim TransRecLen As Integer, POTransFile As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim TCnt As Long, NumTrans As Long, cnt As Integer, Prev As Long
  Dim BadTran As Integer, BadCredits As Double, BadDebits As Double
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, AcctRecNum As Integer
  Dim CntA As Integer, AcctNum As String, LookFor As String
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, IdxAry(CntA)
  Next
    Close AcctIdxFileNum

   'QPrintRC "Relink Purchase Order Transaction Database", 5, 10, 15
   'PrintHelp "System Operations.  Please wait."
  OpenAcctFile GLAcctFile, NumAccts
  TransRecLen = Len(POTrans)
  POTransFile = FreeFile
  Open "POTRANS.DAT" For Random Access Read Write Shared As POTransFile Len = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen
  If NumTrans& > 32767 Then
    Close
    MsgBox "TOO MANY PO TRANSACTIONS, MUST PURGE FIRST.", vbOKOnly, "ERROR"
    Call MainLog("Too Many PO Trans- can't relink")
    Exit Sub
  End If
   'LOCK POTransFile
   'LOCK GLAcctFile
  FrmShowPctComp.Label1 = "Initialize Transactions"
  FrmShowPctComp.Show , x
  DeActivateControls x
  DoEvents

   LogFile = FreeFile
   LogFileName$ = "GLLINK.LOG"
   Open LogFileName$ For Append As #LogFile

   Print #LogFile,
   Print #LogFile, "Accounting Database relink started @ " + Date$ + " @ "; Time$

   '-Set the pointers in the transaction file to zero
   For TCnt& = 1 To NumTrans&
     FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
     Get POTransFile, TCnt&, POTrans
     POTrans.NextTran = 0
     Put POTransFile, TCnt&, POTrans
   Next

  FrmShowPctComp.Label1 = "Initialize Accounts"
  FrmShowPctComp.Show , x

   '-Set the po pointers in the account file to zero
   For cnt = 1 To NumAccts
     FrmShowPctComp.ShowPctComp cnt, NumAccts
      Get GLAcctFile, cnt, Acct
      Acct.FrstPTran = 0
      Acct.Encumb = 0
      Acct.LastPTran = 0
      Put GLAcctFile, cnt, Acct
   Next
  
  FrmShowPctComp.Label1 = "Relink Transactions"
  FrmShowPctComp.Show , x
   '-Start the relink process
   For TCnt& = 1& To NumTrans&
    
      '-Something to look at while this is going on
    FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
    

      Get POTransFile, TCnt&, POTrans

      '-Find the record number of the account
      'AcctRecNum = FindAcct(POTrans.AcctNum)
''''-Find the record number of the account
      AcctNum$ = Trim$(POTrans.AcctNum)
'****Make The Find Faster!!!!
      
      AcctRecNum = 0
'-Find the record number of the account
      For CntA = 1 To NumAIdxRecs
        'Here you put Jump Around Code To Speed UP MOre!!!
        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
        If AcctNum$ = LookFor$ Then
          
          AcctRecNum = IdxAry(CntA).RecNum
          'AcctRecNum = CntA
          Exit For
        Else
          AcctRecNum = 0
        End If
      Next
      '-If we find the account
      If AcctRecNum > 0 Then
         Get GLAcctFile, AcctRecNum, Acct
        '''' If AcctNum$ = "11-426-73" Then Stop
         '--update running encumbrance balance here
         Select Case Acct.Typ
          Case "E", "A"
            Acct.Encumb = Round#(Acct.Encumb + POTrans.DrAmt - POTrans.CrAmt)
          Case "L", "R"
            Acct.Encumb = Round#(Acct.Encumb + POTrans.CrAmt - POTrans.DrAmt)
         End Select
         Put GLAcctFile, AcctRecNum, Acct

         '-Check out the pointer to the first transaction
         Select Case Acct.FrstPTran

           '-If this is the first transaction for this account
           Case 0
               '-Set first and last pointers to this transaction
               Acct.FrstPTran = TCnt&
               Acct.LastPTran = TCnt&
               Put GLAcctFile, AcctRecNum, Acct

            '-If there are already transactions for this account
            Case Is > 0
               '-Remember the pointer to the last transaction.
               Prev& = Acct.LastPTran

               '-Set the last trans pointer to this transaction
               Acct.LastPTran = TCnt&
               Put GLAcctFile, AcctRecNum, Acct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               Get POTransFile, Prev&, POTrans
               POTrans.NextTran = TCnt&
               Put POTransFile, Prev&, POTrans

         End Select

      Else  '-could not find the account
         BadTran = BadTran + 1
         'Print Using; "Orphaned transactions: #####"; BadTran

         '-Keep a list of orphaned transactions.
         GoSub LogPO

      End If

   Next

   '-we're done here
   'UNLOCK POTransFile
   'UNLOCK GLAcctFile
''
''
FixPOEncumb
   If BadTran > 0 Then
      '-Errors in trans file
      Print #LogFile,
      Print #LogFile, "Orphan Transaction Totals:";
      Print #LogFile, Tab(58); Using("#,###,###.##", Str$(BadDebits#));
      Print #LogFile, Tab(70); Using("#,###,###.##", Str$(BadCredits#))
      Print #LogFile, "Relink completed @ " + Date$ + " @ " + Time$
      Print #LogFile, "Orphan transactions encountered! Call Customer Support."
      Call MainLog("RelinkPO Orphans.")
   Else
      '-No errors in trans file
      Print #LogFile, "Relink of PO Transaction successful. " + Date$ + " @ " + Time$
   End If

   Close
   If NumTrans& = 0 Then
      FrmShowPctComp.ShowPctComp 1, 1
   End If

   ActivateControls x
   '-Tell user we're done.
   If BadTran > 0 Then
      '-Errors in trans file
      If Not NM Then
      MsgBox "View Error Log.", vbOKOnly, "Errors Found"
      ViewPrint LogFileName$, "Link Log"
      End If
      Call MainLog("Error Relinking POs")
   Else
      '-No errors in trans file
      If Not NM Then
      MsgBox "Purchase Order transaction relink successful.", vbOKOnly, "Compete"
      End If
      Call MainLog("PO Relink Successful")
   End If
  
Exit Sub

LogPO:
   ToPrint$ = Space$(132)
   LSet ToPrint$ = POTrans.AcctNum
   Mid$(ToPrint$, 18) = Format(DateAdd("d", (POTrans.TRDATE), "12-31-1979"), "mm/dd/yy")
   Mid$(ToPrint$, 30) = Left$(POTrans.Desc, 15)
   Mid$(ToPrint$, 46) = POTrans.Ref
   Mid$(ToPrint$, 58) = Using("#,###,###.##", Str$(POTrans.DrAmt))
   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(POTrans.CrAmt))
   Mid$(ToPrint$, 85) = "Record:" + Str$(TCnt&)
   Print #LogFile, ToPrint$
   BadDebits# = BadDebits# + POTrans.DrAmt
   BadCredits# = BadCredits# + POTrans.CrAmt
Return

End Sub
  
Private Sub FixPOEncumb()   'This is to fix the Encumberance recalc in reindex PO's
  Dim MaxPO As Integer, APLRecLen As Integer
  Dim APLedgerFile As Integer, NumTran As Long, APDRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VendorFile As Integer
  Dim NumVRecs As Integer, AcctFileNum As Integer
  Dim OhShoot As Boolean, cnt As Long, NumPo As Integer, Dept As String
  Dim Rec As Long, NumGLAcctRecs As Integer, Cnt1 As Integer
  Dim Encumb As Double, UnEncumb As Double
  Dim TotEnc As Double, TotUn As Double, NextDist As Long, DistAmt As Double
  Dim Found As Boolean, Amt As Double, DistAcctRec As Integer
  Dim Vendor As VendorRecType
  Dim Acct As GLAcctRecType
  'MaxPO = 400
  'ReDim POList(1 To MaxPO) As GLAcctIndexType   '--borrowing this type
  ReDim POList(1 To 1) As ChkSortType    'use this for long recnum
  '--Get a list of active funds
  Dim ApLedger As APLedger81RecType
  APLRecLen = Len(ApLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  OhShoot = False
  For cnt = 1 To NumTran&
   ' Pct$ = Str$(Int((cnt / NumTran&) * 100))
   ' QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1

    Get APLedgerFile, cnt, ApLedger
    If ApLedger.TRCode = 4 And ApLedger.TRCode <> -4 Then
        NumPo = NumPo + 1
        ReDim Preserve POList(1 To NumPo) As ChkSortType    'use this for long recnum
'        If NumPo = MaxPO Then
'          OhShoot = True
'          Exit For
'        Else
          POList(NumPo).Record = cnt
          POList(NumPo).CHKinfo = Left$(ApLedger.PONum, 14)
'        End If
    End If
  Next

  If OhShoot = True Then
    Close
    MsgBox "Error: Available elements exceed needs. Unable to run report.", vbOKOnly, "Error"
    Exit Sub
  End If

'  If NumPO > 0 Then
'
'    If ShowEnc Then
'      GoSub ClearEnc
'    End If
'
'  Else
  If NumPo > 0 Then
    GoSub ClearEnc
    GoSub SortPO
    GoSub PrintPOList
  Else
    GoSub cancelthis
  End If

  

  Exit Sub


SortPO:
  ReDim Preserve POList(1 To NumPo) As ChkSortType
  Dim lngCurLow As Long, lngCurHigh As Long
  lngCurLow = LBound(POList)
  lngCurHigh = UBound(POList)
  QPOSort POList(), lngCurLow, lngCurHigh
  Return

PrintPOList:
  For cnt = 1 To NumPo

    Rec = POList(cnt).Record
    Get APLedgerFile, Rec, ApLedger
    Get VendorFile, ApLedger.VRecNum, Vendor
'    'IF APLedger.Amt < 100000 THEN
'    '  'STOP
'    '  APLedger.Amt = 0
'    'END IF
    

      NextDist& = ApLedger.FrstDist
      DistAmt# = 0

      If NextDist& > 0 Then

        Do
          Get APDistFile, NextDist&, APDist

          DistAmt# = DistAmt# + APDist.DistAmt
          NextDist& = APDist.NextDist

          If APDist.DistStat <> "L" And APDist.DistStat <> "T" Then
            GoSub UpdateGLAcct
          End If
       Loop Until NextDist& = 0
      End If
  Next
Return

ClearEnc:
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  For Cnt1 = 1 To NumGLAcctRecs
    Get AcctFileNum, Cnt1, Acct
    Acct.Encumb = 0
    Put AcctFileNum, Cnt1, Acct
  Next Cnt1
  Close AcctFileNum
  Return

UpdateGLAcct:   'Reseting the Encumbered Amt
  Amt# = APDist.DistAmt
  DistAcctRec = AcctFind(APDist.DistAcctNum)
  If DistAcctRec > 0 Then
    OpenAcctFile AcctFileNum, NumGLAcctRecs
    Get AcctFileNum, DistAcctRec, Acct
    Acct.Encumb = Acct.Encumb + Amt#
    Put AcctFileNum, DistAcctRec, Acct
    Close AcctFileNum
  End If
 Return
 
cancelthis:
  Exit Sub
End Sub
Public Sub QPOSort(Idxbuff() As ChkSortType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As ChkSortType
  Dim Temp2 As ChkSortType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).CHKinfo < Temp.CHKinfo
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.CHKinfo < Idxbuff(lngCurHigh).CHKinfo
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QPOSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QPOSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub ReLinkTrans(formname As Form, Optional NM As Boolean)
  Dim First As Long, Last As Long, RecNo As Long, AcctRecNum As Integer
  Dim DrAmt As Double, CrAmt As Double, RCnt As Long, BadTran As Integer
  Dim Diff As Double, Bal As Double, ToPrint1 As String, ToPrint3 As String
  Dim CommaFmt As String, TotalFmt As String, ToPrint2 As String
  Dim GLTransFile As Integer, NumTrans As Long, cnt As Integer
  Dim GLAcctFile As Integer, NumAccts As Integer, Prev As Long
  Dim LogFile As Integer, LogFileName As String, TCnt As Long
  Dim BadDebits As Double, BadCredits As Double, ToPrint As String
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim CntA As Integer, AcctNum As String, LookFor As String
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
  For CntA = 1 To NumAIdxRecs
    Get AcctIdxFileNum, CntA, IdxAry(CntA)
  Next
    Close AcctIdxFileNum
   OpenTransFile GLTransFile, NumTrans&
   OpenAcctFile GLAcctFile, NumAccts

   Lock GLTransFile
   Lock GLAcctFile

   LogFile = FreeFile
   LogFileName$ = "GLLINK.LOG"
   Open LogFileName$ For Append As #LogFile
   FrmShowPctComp.Label1 = "Initializing transaction file."
   FrmShowPctComp.CmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
   DoEvents
'''   EnableCloseButton Me.hwnd, False
'''   Me.cmdExit.Enabled = False
'''   Me.cmdGo.Enabled = False


   '-Set the pointers in the transaction file to zero
   For TCnt& = 1 To NumTrans&
      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&

      Get GLTransFile, TCnt&, GLTrans
      GLTrans.NextTran = 0
      Put GLTransFile, TCnt&, GLTrans
  Next          'Process next transaction

   FrmShowPctComp.Label1 = "Initializing account file."
   FrmShowPctComp.CmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
   DoEvents

   '-Set the pointers in the account file to zero
   For cnt = 1 To NumAccts
      FrmShowPctComp.ShowPctComp cnt, NumAccts
      Get GLAcctFile, cnt, GLAcct
      GLAcct.FrstTran = 0
      GLAcct.Bal = 0
      GLAcct.LastTran = 0
      Put GLAcctFile, cnt, GLAcct
  Next          'Process next transaction
   FrmShowPctComp.Label1 = "Relinking."
   FrmShowPctComp.CmdCancel.Enabled = False
   FrmShowPctComp.Show , formname
   DoEvents

    '-Start the relink process
   For TCnt& = 1& To NumTrans&
      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&

      '-Something to look at while this is going on

      Get GLTransFile, TCnt&, GLTrans
      AcctNum$ = Trim$(GLTrans.AcctNum)
'****Make The Find Faster!!!!

'-Find the record number of the account
      For CntA = 1 To NumAIdxRecs
        'Here you put Jump Around Code To Speed UP MOre!!!
        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
        If AcctNum$ = LookFor$ Then
          
          AcctRecNum = IdxAry(CntA).RecNum
          'AcctRecNum = CntA
        Exit For
        End If
      Next

      '-If we find the account
      If AcctRecNum > 0 Then
         Get GLAcctFile, AcctRecNum, GLAcct

         '-Check out the pointer to the first transaction
         Select Case GLAcct.FrstTran

           '-If this is the first transaction for this account
           Case 0
               '-Set first and last pointers to this transaction
               GLAcct.FrstTran = TCnt&
               GLAcct.LastTran = TCnt&
               Put GLAcctFile, AcctRecNum, GLAcct

            '-If there are already transactions for this account
            Case Is > 0
               '-Remember the pointer to the last transaction.
               Prev& = GLAcct.LastTran

               '-Set the last trans pointer to this transaction
               GLAcct.LastTran = TCnt&
               Put GLAcctFile, AcctRecNum, GLAcct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               Get GLTransFile, Prev&, GLTrans
               GLTrans.NextTran = TCnt&
               Put GLTransFile, Prev&, GLTrans

               'update running balance here

         End Select

      Else  '-could not find the account
         BadTran = BadTran + 1

         'Trans.Marked = -1
         'PUT GLTransFile, TCnt&, Trans
         'Trans.Marked = 0

         '-Keep a list of orphaned transactions.
         GoSub Logit

      End If

   Next
'''  Me.cmdExit.Enabled = True
'''  Me.cmdGo.Enabled = True
'''  EnableCloseButton Me.hwnd, True

   '-we're done here
   Unlock GLTransFile
   Unlock GLAcctFile

   If BadTran > 0 Then
      '-Errors in trans file
      Print #LogFile,
      Print #LogFile, "Orphan Transaction Totals:";
      Print #LogFile, Tab(58); Using("#,###,###.##", Str$(BadDebits#))
      Print #LogFile, Tab(70); Using("#,###,###.##", Str$(BadCredits#))
      Print #LogFile, "Relink completed @ " + Date$ + " @ " + Time$
      Print #LogFile, "Orphan transactions encountered! Call Customer Support."
      Call MainLog("RelinkGL Orphans.")
   Else
      '-No errors in trans file
      Call MainLog("RelinkGL successful.")
      If Not NM Then
      MsgBox "Relink of Accounting Databases successful. " + Date$ + "@" + Time$, vbOKOnly, "Relink Successful"
      End If
   End If

   Close

   '-Tell user we're done.
   If BadTran > 0 Then
      '-Errors in trans file
      
      If MsgBox("Errors Encountered, Select Ok to view log or Cancel to Exit.", vbOKCancel, "Error Log") = vbOK Then
        ViewPrint LogFileName$, "Error Log"
      End If
   End If

Exit Sub

Logit:
   ToPrint$ = Space$(132)
   LSet ToPrint$ = GLTrans.AcctNum
   Mid$(ToPrint$, 18) = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
   Mid$(ToPrint$, 30) = Left$(GLTrans.Desc, 15)
   Mid$(ToPrint$, 46) = GLTrans.Ref
   Mid$(ToPrint$, 58) = Using("#'###'###.##", Str$(GLTrans.DrAmt))
   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(GLTrans.CrAmt))
   Mid$(ToPrint$, 85) = "Record:" + Str$(TCnt&)
   Print #LogFile, ToPrint$
   BadDebits# = BadDebits# + GLTrans.DrAmt
   BadCredits# = BadCredits# + GLTrans.CrAmt
Return


End Sub
Public Sub RelinkBgtTrans(formname As Form, Optional NM As Boolean)
  Dim BTrans As GLTransRecType
  Dim Acct As GLAcctRecType
  Dim TransRecLen As Integer, BgtTransFile As Integer, NumTrans As Long
  Dim GLAcctFile As Integer, NumAccts As Integer, TCnt As Long
  Dim LogFile As Integer, LogFileName As String, cnt As Integer
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim CntA As Integer, AcctNum As String, LookFor As String
  Dim Prev As Long, AcctRecNum As Integer, BadTran As Integer
  Dim ToPrint As String, TBGT As Double
'  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'  ReDim IdxAry(1 To NumAIdxRecs) As GLAcctIndexType
'  For CntA = 1 To NumAIdxRecs
'    Get AcctIdxFileNum, CntA, IdxAry(CntA)
'  Next
'    Close AcctIdxFileNum


   TransRecLen = Len(BTrans)
   BgtTransFile = FreeFile
   Open "BGTTRANS.DAT" For Random As BgtTransFile Len = TransRecLen
   NumTrans& = LOF(BgtTransFile) \ TransRecLen

   OpenAcctFile GLAcctFile, NumAccts

   Lock BgtTransFile
   Lock GLAcctFile

   LogFile = FreeFile
   LogFileName$ = "GLLINK.LOG"
   Open LogFileName$ For Append As #LogFile
   Print #LogFile,
   Print #LogFile, "Budget Database relink started @ " + Date$ + " @ "; Time$
   FrmShowPctComp.Label1 = "Initializing Account Transactions."
   FrmShowPctComp.Show , formname
   DoEvents

   '-Set the pointers in the transaction file to zero
   For TCnt& = 1 To NumTrans&
      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
      Get BgtTransFile, TCnt&, BTrans
      BTrans.NextTran = 0
      Put BgtTransFile, TCnt&, BTrans
   Next

   FrmShowPctComp.Label1 = "Initializing Budget Transactions."
   FrmShowPctComp.Show , formname
   DoEvents

'   -Set the budget pointers in the account file to zero
   For cnt = 1 To NumAccts
      FrmShowPctComp.ShowPctComp cnt, NumAccts
      Get GLAcctFile, cnt, Acct
      Acct.FrstBTran = 0
      Acct.Bgt = 0
      Acct.LastBTran = 0
      Put GLAcctFile, cnt, Acct
   Next
   FrmShowPctComp.Label1 = "Relinking Budget Transactions."
   FrmShowPctComp.Show , formname
   DoEvents

   '-Start the relink process
   For TCnt& = 1& To NumTrans&

      '-Something to look at while this is going on
     
      FrmShowPctComp.ShowPctComp TCnt&, NumTrans&
      Get BgtTransFile, TCnt&, BTrans

      'AcctNum$ = Trim$(BTrans.AcctNum)
'-Find the record number of the account
'      For CntA = 1 To NumAIdxRecs
'        'Here you put Jump Around Code To Speed UP MOre!!!
'        LookFor$ = Trim$(IdxAry(CntA).AcctNum)
'        If AcctNum$ = LookFor$ Then
'            'AcctRecNum = CntA
'           AcctRecNum = IdxAry(CntA).RecNum
'
'          Exit For
'        End If
'      Next
      AcctRecNum = AcctFind(Trim$(BTrans.AcctNum))
      '-If we find the account
      If AcctRecNum > 0 Then
         Get GLAcctFile, AcctRecNum, Acct

         '-Check out the pointer to the first transaction
         Select Case Acct.FrstBTran

           '-If this is the first transaction for this account
           Case 0
               '-Set first and last pointers to this transaction
               Acct.FrstBTran = TCnt&
               Acct.LastBTran = TCnt&
               Put GLAcctFile, AcctRecNum, Acct

            Case Is > 0  '-There are already transactions for this account
               '-Remember the pointer to the last transaction.
               Prev& = Acct.LastBTran
               '-Set the last trans pointer to this transaction
               Acct.LastBTran = TCnt&
               Put GLAcctFile, AcctRecNum, Acct

               '-Get the last previous transaction and set its
               '-next tran pointer to this transaction
               Get BgtTransFile, Prev&, BTrans
               BTrans.NextTran = TCnt&
               Put BgtTransFile, Prev&, BTrans
            Case Else
         End Select
         ' TBGT = 0
         '--update the Acct's Budget Balance
         Get BgtTransFile, TCnt&, BTrans
         Select Case Acct.Typ
         
            Case "A", "E"
               'Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.DrAmt) - Round#(BTrans.CrAmt)
                Acct.Bgt = Round#(Acct.Bgt + BTrans.DrAmt - BTrans.CrAmt)
               ' TBGT = Round(BTrans.DrAmt - BTrans.CrAmt)
            Case "L", "R"
               'Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.CrAmt) - Round#(BTrans.DrAmt)
                Acct.Bgt = Round#(Acct.Bgt + BTrans.CrAmt - BTrans.DrAmt)
               ' TBGT = Round#(BTrans.CrAmt - BTrans.DrAmt)
         End Select
''         If Acct.Bgt = 0 Then
''           Acct.Bgt = TBGT
''         Else
''           Acct.Bgt = Round#(Acct.Bgt + TBGT)
''         End If
         Put GLAcctFile, AcctRecNum, Acct

      Else  '-could not find the account
         BadTran = BadTran + 1

         'MsgBox "Orphaned transactions: " & Using("#####", BadTran), vbOKOnly, "Errors Found"
         GoSub LogBgtTrans '-Keep a list of orphaned transactions.

      End If
   Next
   If NumTrans& < 1 Then
      FrmShowPctComp.ShowPctComp 1, 1
   End If

   '-we're done
   Unlock BgtTransFile
   Unlock GLAcctFile
   
   '-Tell user we're done.
   If BadTran > 0 Then
      '-Errors in trans file
      Print #LogFile, "Relink encountered ophans. Completed @ " + Date$ + " @" + Time$
      Call MainLog("RelinkBgt Orphans.")
   Else
      '-No errors in trans file
      Print #LogFile, "Relink of Budget Database successful. " + Date$ + " @ " + Time$
      Call MainLog("RelinkBgt Successful.")
      If Not NM Then
      MsgBox "Re-link successful.", vbOKOnly, "Procedure Complete"
      End If
   End If

   Close

Exit Sub

LogBgtTrans:
   ToPrint$ = Space$(132)
   LSet ToPrint$ = BTrans.AcctNum
   Mid$(ToPrint$, 18) = Format(DateAdd("d", BTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
   Mid$(ToPrint$, 30) = Left$(BTrans.Desc, 15)
   Mid$(ToPrint$, 50) = BTrans.Ref
   Mid$(ToPrint$, 60) = Using("#,###,###.##", Str$(BTrans.DrAmt))
   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(BTrans.CrAmt))
   Mid$(ToPrint$, 80) = "Record:" + Str$(TCnt&)
   Print #LogFile, ToPrint$
Return


End Sub
Public Sub FixPOEncumbRpt(theenddate As Integer, FYStartDate)   'This is to fix the Encumberance recalc in reindex PO's
  Dim MaxPO As Integer, APLRecLen As Integer
  Dim APLedgerFile As Integer, NumTran As Long, APDRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, VendorFile As Integer
  Dim NumVRecs As Integer, AcctFileNum As Integer
  Dim OhShoot As Boolean, cnt As Long, NumPo As Integer, Dept As String
  Dim Rec As Long, NumGLAcctRecs As Integer, Cnt1 As Integer
  Dim Encumb As Double, UnEncumb As Double, EndDate As Integer
  Dim TotEnc As Double, TotUn As Double, NextDist As Long, DistAmt As Double
  Dim Found As Boolean, Amt As Double, DistAcctRec As Integer
  Dim Vendor As VendorRecType
  Dim Acct As GLAcctRecType

  'MaxPO = 400
  'ReDim POList(1 To MaxPO) As GLAcctIndexType   '--borrowing this type
  ReDim POList(1 To 1) As ChkSortType    'use this for long recnum
  '--Get a list of active funds
  Dim ApLedger As APLedger81RecType
  APLRecLen = Len(ApLedger)
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
  EndDate = theenddate

  Dim APDist As APDistRecType
  APDRecLen = Len(APDist)
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  OpenVendorFile VendorFile, NumVRecs

  OhShoot = False
  For cnt = 1 To NumTran&
   ' Pct$ = Str$(Int((cnt / NumTran&) * 100))
   ' QPrintRC "Reading..." + Pct$ + "%", 25, 2, -1

    Get APLedgerFile, cnt, ApLedger
    If ApLedger.TRCode = 4 And ApLedger.TRCode <> -4 Then
      If ApLedger.TRDATE >= FYStartDate And ApLedger.TRDATE <= EndDate Then
        NumPo = NumPo + 1
        ReDim Preserve POList(1 To NumPo) As ChkSortType    'use this for long recnum
'        If NumPo = MaxPO Then
'          OhShoot = True
'          Exit For
'        Else
          POList(NumPo).Record = cnt
          POList(NumPo).CHKinfo = Left$(ApLedger.PONum, 14)
'        End If
      End If
    End If
  Next

  If OhShoot = True Then
    Close
    MsgBox "Error: Available elements exceed needs. Unable to run report.", vbOKOnly, "Error"
    Exit Sub
  End If

'  If NumPO > 0 Then
'
'    If ShowEnc Then
'      GoSub ClearEnc
'    End If
'
'  Else
  If NumPo > 0 Then
    GoSub ClearEnc
    GoSub SortPO
    GoSub PrintPOList
  Else
    GoSub ClearEnc
    GoSub cancelthis
  End If

  

  Exit Sub


SortPO:
  ReDim Preserve POList(1 To NumPo) As ChkSortType
  Dim lngCurLow As Long, lngCurHigh As Long
  lngCurLow = LBound(POList)
  lngCurHigh = UBound(POList)
  QPOSort POList(), lngCurLow, lngCurHigh
  Return

PrintPOList:
  For cnt = 1 To NumPo

    Rec = POList(cnt).Record
    Get APLedgerFile, Rec, ApLedger
    Get VendorFile, ApLedger.VRecNum, Vendor
'    'IF APLedger.Amt < 100000 THEN
'    '  'STOP
'    '  APLedger.Amt = 0
'    'END IF
    
      If ApLedger.TRDATE <= EndDate Then
      NextDist& = ApLedger.FrstDist
      DistAmt# = 0

      If NextDist& > 0 Then

        Do
          Get APDistFile, NextDist&, APDist

          DistAmt# = DistAmt# + APDist.DistAmt
          NextDist& = APDist.NextDist

          If APDist.DistStat <> "L" And APDist.DistStat <> "T" Then
            GoSub UpdateGLAcct
          End If
       Loop Until NextDist& = 0
      End If
      End If
  Next
Return

ClearEnc:
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  For Cnt1 = 1 To NumGLAcctRecs
    Get AcctFileNum, Cnt1, Acct
    Acct.Encumb = 0
    Put AcctFileNum, Cnt1, Acct
  Next Cnt1
  Close AcctFileNum
  Return

UpdateGLAcct:   'Reseting the Encumbered Amt
  Amt# = APDist.DistAmt
  DistAcctRec = AcctFind(APDist.DistAcctNum)
  If DistAcctRec > 0 Then
    OpenAcctFile AcctFileNum, NumGLAcctRecs
    Get AcctFileNum, DistAcctRec, Acct
    Acct.Encumb = Acct.Encumb + Amt#
    Put AcctFileNum, DistAcctRec, Acct
    Close AcctFileNum
  End If
 Return
 
cancelthis:
  Exit Sub
End Sub
Public Sub AddrQSort(Idxbuff() As UBServiceAddressIndexType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As UBServiceAddressIndexType
  Dim Temp2 As UBServiceAddressIndexType
  'temp.SearchName
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  'Stop
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).ServiceAddress < Temp.ServiceAddress
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.ServiceAddress < Idxbuff(lngCurHigh).ServiceAddress
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = Idxbuff(lngCurLow)
        Idxbuff(lngCurLow) = Idxbuff(lngCurHigh)
        Idxbuff(lngCurHigh) = Temp2
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      AddrQSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      AddrQSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub SortServiceAddrs(formname As Form)
  Dim CustRecLen As Integer, NumCustRecs As Long, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Long, lngCurLow As Long, lngCurHigh As Long
  Dim IHandle As Integer, IndexName As String, CRec As Long
  'ShowProcessingScrn "Creating " + IndexText$ + " Index"
 ' QPrintRC "    Reading Customer Records     ", 11, 25, -1

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumCustRecs = GetNumOfCust&

  ReDim ServIndex(1 To NumCustRecs) As UBServiceAddressIndexType
  IndexRecLen = Len(ServIndex(1))

  CHandle = FreeFile
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For cnt = 1 To NumCustRecs
    Get CHandle, cnt, UBCustRec(1)
    ServIndex(cnt).ServiceAddress = UBCustRec(1).ServAddr
    ServIndex(cnt).RecNum = cnt
    'ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next

  Close CHandle

  'QPrintRC "         Sorting Index.        ", 11, 25, -1
  lngCurLow = LBound(ServIndex)
  lngCurHigh = UBound(ServIndex)
  AddrQSort ServIndex(), lngCurLow, lngCurHigh
  'SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
 ' QPrintRC "      Writing Index Records      ", 11, 25, -1
  IndexName$ = "UBTEMP.IDX"
  KillFile IndexName$
  IHandle = FreeFile
  'FCreate IndexName$
  Open IndexName$ For Random Shared As IHandle Len = 4
  For cnt = 1 To NumCustRecs
    CRec& = ServIndex(cnt).RecNum
    Put IHandle, cnt, CRec&
    'ShowPctComp cnt, NumCustRecs                'show user percentage complet
  Next
  Close IHandle

  Erase UBCustRec, ServIndex
End Sub
Public Function GetNumOfCust&()
  Dim UBCustFile As String
  ReDim TCustRec(1) As NewUBCustRecType
  Dim RecLen As Integer
  RecLen = Len(TCustRec(1))
  GetNumOfCust = FileSize(UBCustFile) \ RecLen
  Erase TCustRec
End Function
Public Function Date2Num%(txtDate$)
  On Error GoTo BadDate2Num
  If Len(QPTrim$(txtDate$)) = 10 Then
    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
  Else
    Date2Num% = -32767
  End If
  Exit Function

BadDate2Num:
  On Error GoTo 0
  Date2Num% = -32767
End Function
Public Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  If intDate% = -32767 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If
  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function


'*****Save this stuff*************
'Instead of using this - Used the fpmemo savefile feature
'''Public Sub CpyRptFile(rptname As String)
'''  Dim FileSystemObject As Object
'''  Dim newrpt As String, newlen As Integer
'''  newlen = (Len(rptname) - 3)
'''  newrpt = Mid$(rptname, 1, newlen) + "txt"
'''  Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
'''  FileSystemObject.CopyFile rptname, newrpt
'''End Sub
'''
'**************************
'Emode means in Edit and on existing transaction
'which of course on a new or blank record Emode is false
'When first load form if have gjedit records starts in Edit mode
'******************************
'''Private Sub cmdDelete_Click()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  If Emode = True Then
'''    If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
'''      OpenGJEditFile GJEditFileNum, NumEdTrans
'''      GJEdit.Deleted = -1
'''      Put GJEditFileNum, RecNum, GJEdit
'''      Close GJEditFile
'''      GJEdit.Deleted = 0
'''      Call cmdNew_Click
'''    Else
'''      txtDate.SetFocus
'''    End If
'''  End If
'''End Sub

'''Private Sub cmdEdit_Click()
'''  If Check4Trans = True Then
'''    frmGJListing.Show 1, frmGenJournalEntry
'''    If Emode = True Then
'''      SetScreen
'''      DisplayTotals
'''      txtDate.SetFocus
'''      cmdDelete.Enabled = True
'''    Else
'''      Call cmdNew_Click
'''    End If
'''  Else
'''    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
'''    txtDate.SetFocus
'''  End If
'''End Sub

''''Private Sub cmdList_Click()
''''  If Check4Trans = True Then
''''    frmGJListing.Show 1, frmGenJournalEntry
''''    If Emode = True Then
''''      SetScreen
''''      DisplayTotals
''''      txtDate.SetFocus
''''    Else
''''      Call cmdNew_Click
''''    End If
''''  Else
''''    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
''''    txtDate.SetFocus
''''  End If
''''End Sub

'''Private Sub cmdNew_Click()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  OpenGJEditFile GJEditFileNum, NumEdTrans
'''  If NumEdTrans > 0 Then
'''    DisplayTotals
'''    RecNum = NumEdTrans + 1
'''  Else
'''    RecNum = 1
'''  End If
'''  Emode = False
'''  SetScreen
'''  fpcboAcctNumNa.ListIndex = -1
'''  'txtAcctName = ""
'''  txtAmount = 0
'''  txtEntryType.ListIndex = -1
'''  txtDesc = ""
'''  txtRefNum = ""
'''  txtDate.SetFocus
'''  cmdDelete.Enabled = False
'''End Sub

'''Private Sub cmdSave_Click()
'''  Dim TempDate As Integer
'''  If Emode = True Then
'''    If Changed = False Then
'''      If MsgBox("This Entry Has Not Been Changed, Would you like to Make a New Entry?", vbYesNo, "Go to New") = vbNo Then
'''        txtDate.SetFocus
'''        Exit Sub
'''      End If
'''    End If
'''  End If
'''  'CheckValDate is in main module to verify date entered w/correct format
'''  If CheckValDate(txtDate) = True Then
'''    TempDate = DateDiff("d", "12/31/1979", txtDate)
'''    If (TempDate < LPDate) Or (TempDate > HPDate) Then
'''      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
'''      Exit Sub
'''    End If
'''  Else
'''    MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
'''    Exit Sub
'''  End If
'''  If fpcboAcctNumNa.ColText <> "" And txtAmount > 0 And txtEntryType.ListIndex <> -1 And txtDesc <> "" And txtRefNum <> "" Then
'''    Call SaveGJEntry
'''    Call NextNew
'''  Else
'''    MsgBox "You May Not Save A Blank Field.", vbOKOnly, "Correct and Retry"
'''    txtDate.SetFocus
''' End If
'''End Sub

'''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'''  Select Case KeyCode
'''    Case vbKeyDown, vbKeyReturn:
'''      SendKeys "{Tab}"
'''      KeyCode = 0
'''    Case vbKeyUp:
'''      SendKeys "+{Tab}"
'''      KeyCode = 0
'''    Case vbKeyEscape:
'''      SendKeys "%X"
'''      KeyCode = 0
'''    Case vbKeyF10:
'''      SendKeys "%S"
'''      KeyCode = 0
'''    Case vbKeyF3:
'''      SendKeys "%D"
'''      KeyCode = 0
'''    Case vbKeyF2:
'''      SendKeys "%N"
'''      KeyCode = 0
'''    Case vbKeyF4:
'''      SendKeys "%E"
'''      KeyCode = 0
'''    Case vbKeyF5:
'''      SendKeys "%L"
'''      KeyCode = 0
'''    Case Else:
'''  End Select
'''End Sub

'''Private Sub Form_Resize()
'''  If Me.WindowState <> vbMinimized Then
'''    Me.Visible = False
'''    Temp_Class.ResizeControls Me
'''    Me.Visible = True
'''    Me.SetFocus
'''  End If
'''End Sub

'''Private Sub cmdExit_Click()
'''  If Changed = False Then
'''    frmGenJournalMenu.Show
'''    Unload frmGenJournalEntry
'''  Else
'''    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
'''      frmGenJournalMenu.Show
'''      Unload frmGenJournalEntry
'''    Else
'''      txtDate.SetFocus
'''    End If
'''  End If
'''End Sub

'''Private Sub fpcboAcctNumNa_GotFocus()
'''  fpcboAcctNumNa.Action = ActionClearSearchBuffer
'''End Sub
'''
'''Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
'''  If KeyCode = vbKeySpace Then
'''    fpcboAcctNumNa.ListDown = True
'''  End If
'''End Sub
'''
'''
'''Private Sub mnuPrnScn_Click()
'''  PrintForm
'''End Sub
'''
'''Private Sub txtEntryType_KeyDown(KeyCode As Integer, Shift As Integer)
'''  If KeyCode = vbKeySpace Then
'''    txtEntryType.ListDown = True
'''  End If
'''
''''  If KeyCode = vbKeyBack Then
''''    fpcboAcctNumNa.ListIndex = 0
''''  End If
'''End Sub

'Remark this and try my way 8-31-01
'Private Function GetNextRec()
'  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'  Dim CurrRec As Integer, NextRec As Integer
'  Dim Found As Integer, Rec As Integer
'  OpenGJEditFile GJEditFileNum, NumEdTrans
'  If NumEdTrans > 0 Then
'    CurrRec = 0: NextRec = 1
'      'RecNum = GetNextRec(GJEditFileNum, NumEdTrans, CurrRec, NextRec)
'    Found = 0
'    Rec = CurrRec
'    Do
'      Rec = Rec + NextRec
'      If Rec > NumRecs Or Rec <= 0 Then
'        Found = 0
'        Exit Do
'      End If
'      Get FileNum, Rec, GJEdit
'      If GJEdit.Deleted = 0 Then
'        Found = 1
'        Exit Do
'      End If
'    Loop
'    If Found = 0 Then
'      RecNum = CurrRec
'    Else
'      RecNum = Rec
'    End If
'    If RecNum = 0 Then
'      Close GJEditFile
'      Kill "GJEdit.DAT"
'      Emode = False
'    Else
'      Emode = True
'      Rec2Form RecNum
'      DisplayTotals
'    End If
'  Else
'    Emode = False
'    RecNum = 1
'    Close GJEditFile
'  End If
'
'End Function
'''Private Sub mnuExit_Click()
'''  Call cmdExit_Click
'''End Sub
'''
'''Private Sub mnuPrint_Click()
''''Printer.Print
'''End Sub
'''
'''Private Sub SaveGJEntry()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  OpenGJEditFile GJEditFileNum, NumEdTrans
'''  GJEdit.Deleted = 0
'''  GJEdit.TRDATE = DateDiff("d", "12/31/1979", txtDate)
'''  fpcboAcctNumNa.Col = 1
'''  GJEdit.AcctNum = fpcboAcctNumNa.ColText
'''  fpcboAcctNumNa.Col = 2
'''  GJEdit.AcctName = fpcboAcctNumNa.ColText
'''  GJEdit.EType = Mid$(txtEntryType.Text, 1, 1)
'''  GJEdit.Desc = Trim(txtDesc)
'''  GJEdit.Ref = Trim(txtRefNum)
'''  If txtEntryType.Text = "Debit" Then
'''    GJEdit.DrAmt = txtAmount.DoubleValue()
'''    GJEdit.CrAmt = 0
'''  Else
'''    GJEdit.CrAmt = txtAmount.DoubleValue()
'''    GJEdit.DrAmt = 0
'''  End If
'''  GJEdit.Src = "GJ" + Format$(Now, "mmddyy")
'''  If Emode = False Then
'''    If NumEdTrans > 0 Then
'''      RecNum = NumEdTrans + 1
'''    Else
'''      RecNum = 1
'''    End If
'''  End If
'''  Put GJEditFileNum, RecNum, GJEdit
'''  Close GJEditFile
'''End Sub
'''
'''Private Function SetScreen()
'''  If Emode = False Then  'This is in New Mode
'''    cmdNew.Enabled = False
'''    cmdEdit.Enabled = True
'''    lblNew.Visible = True
'''    lblEdit.Visible = False
'''  Else               'This is in Edit Mode
'''    cmdNew.Enabled = True
'''    cmdEdit.Enabled = False
'''    lblNew.Visible = False
'''    lblEdit.Visible = True
'''  End If
'''End Function
'''Private Sub DisplayTotals()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  Dim cnt As Integer, TotDr As Double, TotCr As Double
'''  OpenGJEditFile GJEditFileNum, NumEdTrans
'''  TotDr = 0: TotCr = 0
'''  For cnt = 1 To NumEdTrans
'''    Get GJEditFileNum, cnt, GJEdit
'''    If Not GJEdit.Deleted Then
'''      TotDr = Round#(TotDr + GJEdit.DrAmt)
'''      TotCr = Round#(TotCr + GJEdit.CrAmt)
'''    End If
'''  Next
'''  Close GJEditFileNum
'''  txtDebits = TotDr
'''  txtCredits = TotCr
'''End Sub
'''Private Function Changed()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  If Emode = False Then
'''    If fpcboAcctNumNa.ListIndex <> -1 Or txtAmount > 0 Then
'''      Changed = True
'''    Else
'''      Changed = False
'''    End If
'''  Else
'''    OpenGJEditFile GJEditFileNum, NumEdTrans
'''    Get GJEditFileNum, RecNum, GJEdit
'''    If txtDate <> Format(DateAdd("d", (GJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy") Then
'''      Changed = True
'''      Close
'''      Exit Function
'''    End If
'''    fpcboAcctNumNa.Col = 1
'''    If fpcboAcctNumNa.ColText <> Trim(GJEdit.AcctNum) Then
'''      Changed = True
'''      Close
'''      Exit Function
'''    End If
'''    If Mid$(txtEntryType.Text, 1, 1) <> GJEdit.EType Then
'''      Changed = True
'''      Close
'''      Exit Function
'''    End If
'''
'''      If txtDesc <> GJEdit.Desc Then
'''        Changed = True
'''        Close
'''        Exit Function
'''
'''
'''      ElseIf txtRefNum <> GJEdit.Ref Then
'''        Changed = True
'''        Close
'''        Exit Function
'''      Else
'''      Changed = False
'''    End If
'''    If GJEdit.EType = "D" Then
'''      If txtAmount <> GJEdit.DrAmt Then
'''        Changed = True
'''        Close
'''        Exit Function
'''      Else
'''        Changed = False
'''      End If
'''    Else
'''      If txtAmount <> GJEdit.CrAmt Then
'''        Changed = True
'''        Close
'''        Exit Function
'''      Else
'''        Changed = False
'''      End If
'''    End If
'''  End If
'''  Close
'''End Function
'
'Private Sub txtAcctNum_Click()
''When select new acct looks up title and displays
'  Dim AcctFile As Integer, NumAccts As Integer, Cnt As Integer, LookFor As String
'  OpenAcctFile AcctFile
'  NumAccts = LOF(AcctFile) / Len(GLAcct)
'  For Cnt = 1 To NumAccts
'    Get AcctFile, Cnt, GLAcct
'      LookFor$ = Trim$(GLAcct.Num)
'      If Trim(txtAcctNum) = LookFor$ Then
'        If GLAcct.Deleted = 0 Then
'          txtAcctName = GLAcct.Title
'          Close AcctFile
'          Exit For
'        End If
'      End If
'  Next
'End Sub
'''Private Sub NextNew()
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  OpenGJEditFile GJEditFileNum, NumEdTrans
'''  If Emode = False Then
'''    If NumEdTrans > 0 Then
'''      DisplayTotals
'''      RecNum = NumEdTrans + 1
'''    Else
'''      RecNum = 1
'''    End If
'''    SetScreen
'''    fpcboAcctNumNa.ListIndex = -1
'''    'txtAcctName = ""
'''    txtAmount = 0
'''    txtEntryType.ListIndex = -1
'''    cmdDelete.Enabled = False
'''    txtDate.SetFocus
'''  Else
'''    Call cmdNew_Click
'''  End If
'''End Sub
'''
'''Private Sub txtDate_LostFocus()
'''  If CheckValDate(txtDate) = False Then
'''    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
'''    txtDate.SetFocus
'''  End If
'''End Sub
'''Private Function Check4Trans()
'''  Dim cnt As Integer, Good As Integer
'''  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
'''  Good = 0
'''  If Exist("Gjedit.dat") Then
'''    OpenGJEditFile GJEditFileNum, NumEdTrans
'''    If NumEdTrans > 0 Then
'''      For cnt = 1 To NumEdTrans
'''        Get GJEditFileNum, cnt, GJEdit
'''        If GJEdit.Deleted = 0 Then
'''          Good = Good + 1
'''        End If
'''      Next
'''    Else
'''      Check4Trans = False
'''    End If
'''  Else
'''    Check4Trans = False
'''  End If
'''  If Good > 0 Then
'''    Check4Trans = True
'''  Else
'''    Check4Trans = False
'''  End If
''' Close GJEditFileNum
''' End Function
''Public Sub ExpHSGLAcct()
''  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
''  Dim AcctFileNum As Integer, NumAccts As Integer
''  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
''  Dim ToPrint As String
''  Dim Header As String, Tempcode As String
''  Dim AcctIdx As GLAcctIndexType
''  Dim GLAcct As GLAcctRecType
''  Dim q As String, ExpFile As String
''   q$ = "|"
''   FrmShowPctComp.Label1 = "Creating GL Account Export"
''   FrmShowPctComp.cmdCancel.Enabled = False
''   FrmShowPctComp.Show , frmGLSetupMenu
''   Header$ = "GLAcctNumber|GLAcctName|AcctType"
''   OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
''   OpenAcctFile AcctFileNum
''   PRNFile = FreeFile
''   ExpFile$ = StartPath + "\Exports\" + "GLAcct.ASC"
''   Open ExpFile$ For Output As #PRNFile
''   Print #PRNFile, Header$
''   For cnt = 1 To NumAIdxRecs
''      Get AcctIdxFileNum, cnt, AcctIdx
''      FrmShowPctComp.ShowPctComp cnt, NumAIdxRecs
''      If FrmShowPctComp.Out = True Then
''        Close
''        Unload FrmShowPctComp
''        GoTo CancelExit
''      End If
''      Get AcctFileNum, AcctIdx.RecNum, GLAcct
''      Howmany = Howmany + 1
''      Print #PRNFile, QPTrim$(GLAcct.Num); q$; QPTrim$(GLAcct.Title); q$; QPTrim$(GLAcct.Typ)
''   Next
''
''CancelExit:
''  Close
''  If Howmany > 0 Then
''    MsgBox "File " & "\Exports\" & ExpFile$ & " Exported with " & Howmany & " GL Accounts.", vbOKOnly, "GL Accounts Exported."
''  Else
''    MsgBox "No Information Found to Export.", vbOKOnly, "This portion Ended"
''  End If
''  Call MainLog("GL Export Accounts, Exported " & Howmany)
''
''End Sub
''Public Sub ExpFunds()
''  Dim FundIdxFileNum As Integer, NumFIdxRecs As Integer
''  Dim FundFileNum As Integer, NumFunds As Integer
''  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
''  Dim Header As String
''  Dim FundIdx As GLFundIndexType
''  Dim Fund As GLFundRecType
''  Dim q As String, ExpFile As String
''   q$ = "|"
''   FrmShowPctComp.Label1 = "Creating GL Fund Export"
''   FrmShowPctComp.cmdCancel.Enabled = False
''   FrmShowPctComp.Show , frmGLSetupMenu
''
''  Header$ = "FundNumber|FundName"
''  OpenFundIdx FundIdxFileNum, NumFIdxRecs
''  OpenFundFile FundFileNum, NumFunds
''  PRNFile = FreeFile
''  ExpFile$ = StartPath + "\Exports\" + "GLFund.ASC"
''  Open ExpFile$ For Output As #PRNFile
''  Print #PRNFile, Header$
''  For cnt = 1 To NumFIdxRecs
''    Get FundIdxFileNum, cnt, FundIdx
''      FrmShowPctComp.ShowPctComp cnt, NumFIdxRecs
''      If FrmShowPctComp.Out = True Then
''        Close
''        Unload FrmShowPctComp
''        GoTo CancelExit
''      End If
''
''    Get FundFileNum, FundIdx.RecNum, Fund
''    Howmany = Howmany + 1
''    Print #PRNFile, QPTrim$(Fund.FundNum); q$; QPTrim$(Fund.Title)
''  Next
''CancelExit:
''  Close
''  If Howmany > 0 Then
''    MsgBox "File " & "\Exports\" & ExpFile$ & " Exported with " & Howmany & " GL Funds.", vbOKOnly, "Funds Exported."
''  Else
''    MsgBox "No Information Found to Export.", vbOKOnly, "This portion Ended"
''  End If
''  Call MainLog("GL Export Funds, Exported " & Howmany)
''End Sub
''Public Sub ExpDepartments()
''  Dim DeptIdxFileNum As Integer, NumDIdxRecs As Integer
''  Dim DeptFileNum As Integer, NumDepts As Integer
''  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
''  Dim Header As String
''  Dim GLDeptIdx As GLDeptIndexType
''  Dim GLDept As GLDeptRecType
''  Dim q As String, ExpFile As String
''   q$ = "|"
''   FrmShowPctComp.Label1 = "Creating GL Dept Export"
''   FrmShowPctComp.cmdCancel.Enabled = False
''   FrmShowPctComp.Show , frmGLSetupMenu
''
''  Header$ = "DeptNumber|DeptName"
''  OpenDeptIdx DeptIdxFileNum, NumDIdxRecs
''  OpenDeptFile DeptFileNum, NumDepts
''  PRNFile = FreeFile
''
''  ExpFile$ = StartPath + "\Exports\" + "GLDept.ASC"
''  Open ExpFile$ For Output As #PRNFile
''  Print #PRNFile, Header$
''  For cnt = 1 To NumDIdxRecs
''    Get DeptIdxFileNum, cnt, GLDeptIdx
''      FrmShowPctComp.ShowPctComp cnt, NumDIdxRecs
''      If FrmShowPctComp.Out = True Then
''        Close
''        Unload FrmShowPctComp
''        GoTo CancelExit
''      End If
''
''    Get DeptFileNum, GLDeptIdx.RecNum, GLDept
''    Howmany = Howmany + 1
''    Print #PRNFile, QPTrim$(GLDept.DeptNum); q$; QPTrim$(GLDept.Title)
''  Next
''CancelExit:
''  Close
''  If Howmany > 0 Then
''    MsgBox "File " & "\Exports\" & ExpFile$ & " Exported with " & Howmany & " GL Depts.", vbOKOnly, "GL Depts Exported."
''  Else
''    MsgBox "No Information Found to Export.", vbOKOnly, "This portion Ended"
''  End If
''  Call MainLog("GL Export Depts, Exported " & Howmany)
''
''End Sub
''Public Sub ExpBankCodes()
''  Dim BankFileNum As Integer, NumBankRecs As Integer
''  Dim PRNFile As Integer, cnt As Integer, Howmany As Integer
''  Dim Header As String
''  Dim GLBank As GLBankRecType
''  Dim q As String, ExpFile As String
''   q$ = "|"
''   FrmShowPctComp.Label1 = "Creating GL Bank Export"
''   FrmShowPctComp.cmdCancel.Enabled = False
''   FrmShowPctComp.Show , frmGLSetupMenu
''   Header$ = "BankNum|BankName|GLAccount"
''
''   OpenBankFile BankFileNum, NumBankRecs
''   PRNFile = FreeFile
''
''  ExpFile$ = StartPath + "\Exports\" + "GLBank.ASC"
''  Open ExpFile$ For Output As #PRNFile
''  Print #PRNFile, Header$
''
''   For cnt = 1 To NumBankRecs
''      Get BankFileNum, cnt, GLBank
''      FrmShowPctComp.ShowPctComp cnt, NumBankRecs
''      If FrmShowPctComp.Out = True Then
''        Close
''        Unload FrmShowPctComp
''        GoTo CancelExit
''      End If
''      If GLBank.Deleted = 0 Then
''
''        Howmany = Howmany + 1
''        Print #PRNFile, QPTrim$(Str$(GLBank.BankNum)); q$; QPTrim$(GLBank.BankName);
''        Print #PRNFile, q$; QPTrim$(GLBank.GLAcct)
''
''      End If
''   Next
''
''CancelExit:
''  Close
''  If Howmany > 0 Then
''    MsgBox "File " & "\Exports\" & ExpFile$ & " Exported with " & Howmany & " GL Banks.", vbOKOnly, "GL Banks Exported."
''  Else
''    MsgBox "No Information Found to Export.", vbOKOnly, "This portion Ended"
''  End If
''  Call MainLog("GL Export Banks, Exported " & Howmany)
''
''End Sub
