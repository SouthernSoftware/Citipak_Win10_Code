Attribute VB_Name = "modGLCommon"
Option Explicit
Dim GLSetup   As GLSetupRecType
Dim GLFund    As GLFundRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLDept    As GLDeptRecType
Dim GLDeptIdx As GLDeptIndexType
Dim GLBank    As GLBankRecType
Dim APInvTax  As APInvTaxRecType
Dim GJEdit    As TrEditRecType
Dim GLTrans   As GLTransRecType
Dim CJEdit    As CJEditRecType
Dim BgtEdit   As TrEditRecType
Dim BgtTrans  As GLTransRecType
Dim OSChk     As OSChkRecType
Dim ApLedger  As APLedger81RecType
Dim APDist    As APDistRecType
Dim apvendor  As VendorRecType


Public screenW As Long
Public coladj As Double



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
      If GLAcct.DELETED = 0 Then
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

Public Function FindFund(FundNum$)
  Dim NumOfFunds As Integer, cnt As Integer, FundFile As Integer
  Dim Match As Boolean, LookFor As String
  FundNum$ = LTrim$(FundNum$)
  OpenFundFile FundFile, NumOfFunds
  'NumOfFunds = LOF(FundFile) / Len(GLFund)
  For cnt = 1 To NumOfFunds
  Get FundFile, cnt, GLFund
  LookFor$ = Trim$(GLFund.FundNum)
  If GLFund.DELETED = 0 Then
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

Public Function SortFundIndex()
  Dim FundIdxFileNum As Integer, NumFFIdxRecs As Integer, FundFileNum As Integer
  Dim NumFunds As Integer, cnt As Integer, GoodFunds As Integer
  Dim OutOfOrder As Boolean, TempIdxRec As GLFundIndexType
  KillFile "GLFund.IDX"
  OpenFundIdx FundIdxFileNum, NumFFIdxRecs
  OpenFundFile FundFileNum, NumFunds
  If NumFunds <= 1 Then    'no need to sort one record
    Close FundIdxFileNum, FundFileNum
    Exit Function
  End If
  ReDim Idxbuff(1 To NumFunds) As GLFundIndexType
  For cnt = 1 To NumFunds
    Get FundFileNum, cnt, GLFund
    If GLFund.DELETED = 0 Then
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
Public Function Exist(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
End Function
Public Sub KillFile(FileName$)
  If Exist(FileName) Then
    Kill FileName$
  End If
End Sub

Public Sub ViewPrint(ReportFile As String, title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmLoadingRpt.Show
   frmViewPrint.ReportName = ReportFile$
   frmViewPrint.Caption = title
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
   frmViewPrint.Show 1
   Unload frmLoadingRpt
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
  Dim I As Integer, J As Integer
  KillFile "GLAcct.IDX"
  FrmShowPctComp.Label1 = "Initializing Account Index."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  If NumAccts <= 1 Then    'no need to sort one record
    Close AcctIdxFileNum, AcctFileNum
    Exit Function
  End If
  ReDim Idxbuff(1 To NumAccts) As GLAcctIndexType
  For CntAc = 1 To NumAccts
    FrmShowPctComp.ShowPctComp CntAc, NumAccts
    Get AcctFileNum, CntAc, GLAcct
    If GLAcct.DELETED = 0 Then
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
  If NumDepts <= 1 Then    'no need to sort one record
    Close DeptIdxFileNum, DeptFileNum
    Exit Function
  End If
  
  ReDim Idxbuff(1 To NumDepts) As GLDeptIndexType
  For cnt = 1 To NumDepts
    Get DeptFileNum, cnt, GLDept
    If GLDept.DELETED = 0 Then
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
      If GLAcct.DELETED = 0 Then
        txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.title) & Chr$(9) & QPStrip(GLAcct.Num)
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
      If GLAcct.DELETED = 0 Then
        'Tried this for one column combo
        'TempList = (QPTrim(GLAcct.Num)) & "   " & Trim(GLAcct.Title)
        TempList = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.title) & Chr$(9) & QPStrip(GLAcct.Num)
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
      If GLFund.DELETED = 0 Then
           '''txtField.InsertRow = Str$(GLFund.FundNum) & Chr$(32) & QPTrim(GLFund.Title)

        txtField.InsertRow = Str$(GLFund.FundNum) & Chr$(9) & QPTrim(GLFund.title)
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
      If GLDept.DELETED = 0 Then
        txtField.InsertRow = Str$(GLDeptIdx.RecNum) & Chr$(9) & QPTrim(GLDept.DeptNum) & Chr$(9) & QPTrim(GLDept.title)
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
  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
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
           GLTrans.DESC = Tran2Post.DESC
           GLTrans.CrAmt = Tran2Post.CrAmt
           GLTrans.DrAmt = Tran2Post.DrAmt
           GLTrans.Ref = Tran2Post.Ref
           GLTrans.Src = Tran2Post.Src
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
   Print #GLLogFile, "Description    :"; Tran2Post.DESC
   Print #GLLogFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #GLLogFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #GLLogFile,
Return
LogTempErr:
   Print #PRNFile, "Unpostable Transaction"
   Print #PRNFile, "Record Number  :"; Str$(cnt)
   Print #PRNFile, "Account Number :"; Tran2Post.AcctNum
   Print #PRNFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #PRNFile, "Description    :"; Tran2Post.DESC
   Print #PRNFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #PRNFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #PRNFile,
Return
ItsBroke:
  BadTrans = BadTrans + 1
  Print #PRNFile, "Error *** Call Software Support***"
  Print #PRNFile, "Record Number :"; Str$(cnt); Tran2Post.AcctNum
  Print #PRNFile, "Error Code"; Str(Err.Number)
  Resume Next
  
End Function
Public Function GetBankList(txtName As fpCombo)
  Dim Bank As GLBankRecType, BankRecLen As Integer, BankFile As Integer
  Dim BankNum As String, cnt As Integer, NumBanks As Integer
  BankRecLen = Len(Bank)
  OpenBankFile BankFile, NumBanks
  If NumBanks = 0 Then
    Close BankFile
    Exit Function
  End If
  txtName.Row = -1
  For cnt = 1 To NumBanks
    Get BankFile, cnt, Bank
    If Bank.DELETED = 0 Then
      txtName.InsertRow = (Bank.BankNum & Chr$(9) & Mid$(Bank.BankName, 1, 25))
      '(Bank.BankNum & "  " & )
      'txtName. (txtName.NewIndex) = (Val(Bank.BankNum))
    Else
    End If
  Next
  Close BankFile
End Function
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
  CashAcct = GLSetUpRec(1).CashAcct
  CDCash = GLSetUpRec(1).CDCash
  CDDue = GLSetUpRec(1).CDDue
  
  Erase GLSetUpRec
End Sub
'
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
  Dim BadCashAcct As Boolean, PRNFileName As String, JEDebits As Double, JECredits As Double
  Dim LineCnt As Integer, IFFile As Integer, NumTrans As Integer
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  Dim Trans2Post As GLTransRecType
  Dim TmpSortTrans As GLTransRecType
  Dim OutOfOrder As Boolean
  Dim CJEdit As CJEditRecType
  Dim Tr2Post As GLTransRecType
  Dim TmpIFFile As String, CJPrnFile As String, CJEditFile As String
  Dim OSChk As OSChkRecType
  Dim OSChkFile As Integer, NumOSChks As Integer, OSChkFileNum As Integer, CHK As Integer
  Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
'the get list of funds on gj main form
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  ReDim TrFundSum(1 To NumFunds) As Double
  CommaFmt$ = "###,###,###,###.##"
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen

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
    If Not CJEdit.DELFLAG Then
      Active = Active + 1
    End If
  Next
  Close CJEditFileNum
'Give options to cancel posting
  If Active = 0 Then
    MsgBox "No Transactions To Post", vbOKOnly, "Post Canceled"
    Exit Function
  End If
    
  If MsgBox("Are You Sure You Wish to Post.", vbOKCancel, "CD Posting") = vbCancel Then
    Exit Function
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
      Exit Function
    End If
    RecordNum = 0
    For cnt = 1 To NumEdTrans              'Assign edit file to trans format
      For Fund = 1 To NumFunds
        TrFundSum#(Fund) = 0
      Next
      Get CJEditFileNum, cnt, CJEdit
      If Not CJEdit.DELFLAG Then
        'If Central Depository use Central Deposi Cash Acct
        If CDActive$ = "Y" Then
        CshAcct$ = CashAcct$
        RecordNum = RecordNum + 1
        Tr2Post.AcctNum = CDCash$
        Tr2Post.TRDATE = CJEdit.TRDATE
        Tr2Post.DESC = CJEdit.DESC
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
        CshAcct$ = GetBankGLAcct(Val(CJEdit.RECCODE))
      End If
      'Add each Distribution to the interface file
      For CntD = 1 To 36
        If CJEdit.Dist(CntD).DACREC > 0 Then
          RecordNum = RecordNum + 1
          Tr2Post.AcctNum = CJEdit.Dist(CntD).DACN
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.DESC = CJEdit.DESC
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
          For Fund = 1 To NumFunds - 1
            FundNum$ = Left$(CJEdit.Dist(CntD).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round(CJEdit.Dist(CntD).DAMT)
              Exit For
            End If
          Next 'Fund
        End If
      Next 'Distribution
    'No More Distributions so now create the "Opposite" entries.
    'One to cash or central dep for each fund
      For Fund = 1 To NumFunds - 1
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
        Tr2Post.DESC = CJEdit.DESC
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
            Tr2Post.DESC = CJEdit.DESC
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
      KillFile TmpIFFile$
      MsgBox "Invalid Cash Account, Posting Aborted. Check Journal Report for Invalid Entry.", vbOKOnly, "Posting Aborted"
      Exit Function
    End If
    Close
    
  
    Call Post2GL(TmpIFFile$, BadTrans, formname, False) 'common post & link sub
    If BadTrans <> 0 Then
      Close
      KillFile TmpIFFile$
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      ViewPrint ReportFile$, "Error Log"
      frmCitiCancel.Show
      Unload formname
      'Need to unload menu but how to reference it?????
      Exit Function
    End If
    If CJType = 2 Then
    'Post disbursements to o/s check file
      OpenOSChkFile OSChkFileNum, NumOSChks
      OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
      For CHK = 1 To NumEdTrans
        Get CJEditFileNum, CHK, CJEdit
        If Not CJEdit.DELFLAG Then
          NumOSChks = NumOSChks + 1
          OSChk.ChkNum = Val(CJEdit.DOCREF)
          OSChk.ChkDate = Left$(Format(DateAdd("d", (CJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy"), 8)
  
          OSChk.DESC = CJEdit.DESC
          OSChk.Amt = CJEdit.Amt
          OSChk.Src = 0 ' Why this ????? code as apcheck
          OSChk.Cleared = 0
          Put OSChkFileNum, NumOSChks, OSChk
        End If
      Next
    End If
    Close
    Call Post2GL(TmpIFFile$, BadTrans, formname, True) 'common post & link sub
    If BadTrans <> 0 Then                  'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      GLLogFileName = "GLlog.dat"
      ReportFile$ = "GLlog.dat"
      ViewPrint ReportFile$, "Posting Log"
    End If
    GoSub PrnPostJournal
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
  ReDim SortTrans(1 To NumTrans) As GLTransRecType
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    SortTrans(cnt) = Trans2Post
  Next
  Close
  '*** What is SortT ??
  Do
    OutOfOrder = False          'assume it's sorted
    For cnt = 1 To NumTrans - 1
      If SortTrans(cnt).AcctNum > SortTrans(cnt + 1).AcctNum Then
        LSet TmpSortTrans = SortTrans(cnt)
        LSet SortTrans(cnt) = SortTrans(cnt + 1)
        LSet SortTrans(cnt + 1) = TmpSortTrans
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
  'The SortT below was from old Program **The Section Above (Per Dale) Replaced it.
  'SortT SortTrans(1), NumTrans, 0, 96, 2,14
  IFFile = FreeFile
  Open TmpIFFile$ For Random As IFFile Len = RecLen
  
  For cnt = 1 To NumTrans
    Put IFFile, cnt, SortTrans(cnt)
  Next
  ToPrint$ = Space$(82)
  PRNFile = FreeFile
  PRNFileName$ = "CJPOST.PRN"
  Open PRNFileName$ For Output As #PRNFile
  GoSub CDJEHeader
  For cnt = 1 To NumTrans
    Get IFFile, cnt, Trans2Post
    JEDebits# = JEDebits# + Round(Trans2Post.DrAmt)
    JECredits# = JECredits# + Round(Trans2Post.CrAmt)
   
    LSet ToPrint$ = ""
    LSet ToPrint$ = Format(DateAdd("d", (Trans2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Mid$(ToPrint$, 13) = Trans2Post.AcctNum
    Mid$(ToPrint$, 29) = Left$(Trans2Post.DESC, 15)
    Mid$(ToPrint$, 42) = Trans2Post.Ref
    Mid$(ToPrint$, 45) = Using$(CommaFmt$, Trans2Post.DrAmt)
    Mid$(ToPrint$, 63) = Using$(CommaFmt$, Trans2Post.CrAmt)
    Print #PRNFile, ToPrint$
    LineCnt = LineCnt + 1
    If LineCnt > 55 Then
      Print #PRNFile, Chr$(12)
      GoSub CDJEHeader
    End If
  Next
  
  Print #PRNFile, ' Blank line
  Print #PRNFile, "Posting Totals:";
  Print #PRNFile, Tab(45); Using$(CommaFmt$, JEDebits#);
  Print #PRNFile, Tab(63); Using$(CommaFmt$, JECredits#);
  Print #PRNFile, Chr$(12)
  Close
  'File Must be closed before going to ViewPrint - Hence the close above
  If CJType = 1 Then
    ViewPrint PRNFileName$, "Cash Receipt Journal Post Report"
  ElseIf CJType = 2 Then
    ViewPrint PRNFileName$, "Cash Disbursement Journal Post Report"
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
  LineCnt = 5
Return
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
      If BgtEdit.DELETED = 0 Then         'Ok if not deleted
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

'****************************************************************************
'Retrieves the fund title from the fund data file.
'****************************************************************************
'
Public Function GetFundTitle(FundRecNum)
  Dim NumFunds As Integer, FundFileNum As Integer
  Dim GLFund As GLFundRecType
   OpenFundFile FundFileNum, NumFunds
   Get FundFileNum, FundFileNum, GLFund
   GetFundTitle = GLFund.title
   Close FundFileNum

End Function

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
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
      If GLAcct.DELETED = 0 Then
        If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
          txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & Trim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.title) & Chr$(9) & QPStrip(GLAcct.Num)
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
      If GLAcct.DELETED = 0 Then
        If GLAcct.Typ = "E" Or GLAcct.Typ = "R" Then
          TempBudList = Str$(GLAcctidx.RecNum) & Chr$(9) & Trim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.title) & Chr$(9) & QPStrip(GLAcct.Num)
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

Public Function Num2Month%(Dt%)
  Dim d As String, m As String
  d$ = Format(DateAdd("d", Dt%, "12-31-1979"), "mm/dd/yyyy")
  m$ = Right$(d$, 2) + Left$(d$, 2)
  Num2Month% = Val(m$)

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
  GetDeptTitle$ = DeptRec.title
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
      If GLFund.DELETED = 0 Then
        x.AddItem Str$(QPTrim(GLFund.FundNum)) & Chr$(9) & QPTrim(GLFund.title) & Chr$(9) & QPTrim(GLFund.FundNum)
      End If
  Next
  Close FundIdxFileNum
  Close FundFileNum
End Function

Public Sub ReLinkTrans(formname As Form)
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
   Else
      '-No errors in trans file
      MsgBox "Relink of Accounting Databases successful. " + Date$ + "@" + Time$, vbOKOnly, "Relink Successful"
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
   Mid$(ToPrint$, 30) = Left$(GLTrans.DESC, 15)
   Mid$(ToPrint$, 46) = GLTrans.Ref
   Mid$(ToPrint$, 58) = Using("#'###'###.##", Str$(GLTrans.DrAmt))
   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(GLTrans.CrAmt))
   Mid$(ToPrint$, 85) = "Record:" + Str$(TCnt&)
   Print #LogFile, ToPrint$
   BadDebits# = BadDebits# + GLTrans.DrAmt
   BadCredits# = BadCredits# + GLTrans.CrAmt
Return


End Sub
Public Sub RelinkBgtTrans(formname As Form)
  Dim BTrans As GLTransRecType
  Dim Acct As GLAcctRecType
  Dim TransRecLen As Integer, BgtTransFile As Integer, NumTrans As Long
  Dim GLAcctFile As Integer, NumAccts As Integer, TCnt As Long
  Dim LogFile As Integer, LogFileName As String, cnt As Integer
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim CntA As Integer, AcctNum As String, LookFor As String
  Dim Prev As Long, AcctRecNum As Integer, BadTran As Integer
  Dim ToPrint As String
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
   '-Start the relink process
   FrmShowPctComp.Label1 = "Relink Budget Transaction Database."
   FrmShowPctComp.Show , formname
   DoEvents

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

         '--update the Acct's Budget Balance
         Select Case Acct.Typ
            Case "A", "E"
               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.DrAmt) - Round#(BTrans.CrAmt)
            Case "L", "R"
               Acct.Bgt = Round#(Acct.Bgt) + Round#(BTrans.CrAmt) - Round#(BTrans.DrAmt)
         End Select
         Put GLAcctFile, AcctRecNum, Acct

      Else  '-could not find the account
         BadTran = BadTran + 1

         'MsgBox "Orphaned transactions: " & Using("#####", BadTran), vbOKOnly, "Errors Found"
         GoSub LogBgtTrans '-Keep a list of orphaned transactions.

      End If
   Next

   '-we're done
   Unlock BgtTransFile
   Unlock GLAcctFile

   '-Tell user we're done.
   If BadTran > 0 Then
      '-Errors in trans file
      Print #LogFile, "Relink encountered ophans. Completed @ " + Date$ + " @" + Time$
   Else
      '-No errors in trans file
      Print #LogFile, "Relink of Budget Database successful. " + Date$ + " @ " + Time$
      MsgBox "Re-link successful.", vbOKOnly, "Procedure Complete"

   End If

   Close

Exit Sub

LogBgtTrans:
   ToPrint$ = Space$(132)
   LSet ToPrint$ = BTrans.AcctNum
   Mid$(ToPrint$, 18) = Format(DateAdd("d", BTrans.TRDATE, "12-31-1979"), "mm/dd/yy")
   Mid$(ToPrint$, 30) = Left$(BTrans.DESC, 15)
   Mid$(ToPrint$, 50) = BTrans.Ref
   Mid$(ToPrint$, 60) = Using("#,###,###.##", Str$(BTrans.DrAmt))
   Mid$(ToPrint$, 70) = Using("#,###,###.##", Str$(BTrans.CrAmt))
   Mid$(ToPrint$, 80) = "Record:" + Str$(TCnt&)
   Print #LogFile, ToPrint$
Return


End Sub

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
'      Kill "GLGJED.DAT"
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
'''  If Exist("Glgjed.dat") Then
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

