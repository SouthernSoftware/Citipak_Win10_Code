Attribute VB_Name = "modCMCommon"
Option Explicit

Public Function OKDepRefund(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBCustF As Integer
  UBCustRecLen = Len(UBCustRec(1))
  If RecNo& > 0 Then
    UBCustF = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
    Get UBCustF, RecNo&, UBCustRec(1)
    Close UBCustF
  
    If UBCustRec(1).DepositAmt <= 0 Then
      'OK = MsgBox%("UB", "NODPOSIT")
      OKDepRefund = False
    Else
      OKDepRefund = True
    End If
    If OKDepRefund = False Then
      frmMsgDialog.RetLabel = "-2"
      FntSize = frmMsgDialog.Label(3).FontSize
      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
      FntSize = frmMsgDialog.Label(1).FontSize
      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
      MsgText(0) = "ERROR!"
      MsgText(1) = ""
      MsgText(2) = "NO DEPOSIT"
      MsgText(3) = ""
      MsgText(4) = "This Account Has NO Deposit on File"
      MsgText(5) = ""
      GetOKorNot MsgText(), True
    End If

  End If
End Function

Public Sub UPDateOK()
  frmDataUpdated.Show vbModal
End Sub
Public Sub OpenSetupFile(SetupFileNum)
  Dim GLSetupRecLen  As Integer
  GLSetupRecLen = Len(GLSetup)
  SetupFileNum = FreeFile
  Open UBPath$ + "GLSetup.DAT" For Random Shared As SetupFileNum Len = GLSetupRecLen
End Sub
Public Sub GetUBBankINfo()
  Dim CMBnkAcct As CMBankAcctRecType
  Dim CMBnkAcctLen As Integer, CMFile As Integer
  On Local Error GoTo ubb
  CMBnkAcctLen = Len(CMBnkAcct)
  CMFile = FreeFile
  Open UBPath + "CMBkAcct.DAT" For Random Shared As CMFile Len = CMBnkAcctLen
  Get CMFile, 1, CMBnkAcct
    BnkAcctNum$ = QPTrim$(CMBnkAcct.COMPACCT)
  Close
  Exit Sub
ubb:
   BnkAcctNum$ = " "
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
Public Sub OpenAcctIdx(AcctIdxFileNum, NumAIdxRecs)
  Dim GLAcctIdxLen As Integer
  GLAcctIdxLen = Len(GLAcctidx)
  AcctIdxFileNum = FreeFile
  Open UBPath$ + "GLAcct.Idx" For Random Shared As AcctIdxFileNum Len = GLAcctIdxLen
  NumAIdxRecs = LOF(AcctIdxFileNum) \ GLAcctIdxLen
End Sub

'!!! Procedures below Needed for reports!!! Mark with!!!
'Make sure to check w/Dale  PS
'!!! Added Round on 4-17-03
Public Function Round#(ByVal n#)
  Round# = (Int(n# * 100 + 0.5000001)) / 100
End Function
'loads Work Order Defaults into fpcombos
Public Sub GetWOList(x As fpCombo)
  Dim cnt As Long, NumWOs As Long
  Dim WorkOrderDefLen As Integer
  Dim UBWrkOrdD As Integer

  Dim WorkOrderDef As WorkOrderDefType
  WorkOrderDefLen = Len(WorkOrderDef)

  UBWrkOrdD = FreeFile
  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
  NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
  For cnt = 1 To NumWOs
    Get UBWrkOrdD, cnt, WorkOrderDef
      If WorkOrderDef.Deleted <> True Then
        x.InsertRow = Str(cnt) & Chr$(9) & QPTrim(WorkOrderDef.WOType)
      End If
  Next
  Close
End Sub
Public Function OKDeleteCust(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, TotalBalance As Double
  Dim M1 As String, M2 As String
  Dim UBCustRecLen As Integer, UBCustF As Integer
  If RecNo& > 0 Then
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCustF = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
  Get UBCustF, RecNo&, UBCustRec(1)
  Close UBCustF

  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  If TotalBalance# <> 0 Then
    UBLog "NODELETE:" + Str$(RecNo&) + " BAL:" + Str$(TotalBalance#)
    M1$ = "This account HAS A BALANCE"
    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
    OKDeleteCust = False
  ElseIf UBCustRec(1).DepositAmt <> 0 Then
    UBLog "NODELETE:" + Str$(RecNo&) + " DEP:" + Str$(UBCustRec(1).DepositAmt)
    M1$ = "This account HAS A DEPOSIT"
    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
    OKDeleteCust = False
  ElseIf UBCustRec(1).Status <> "I" Then
    UBLog "NODELETE:" + Str$(RecNo&) + " NOT INACTIVE"
    M1$ = "This account IS NOT INACTIVE"
    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
    OKDeleteCust = False
  Else
    OKDeleteCust = True
  End If
  If OKDeleteCust = False Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR!"
    MsgText(1) = ""
    MsgText(2) = M1$
    MsgText(3) = ""
    MsgText(4) = M2$
    MsgText(5) = ""
    GetOKorNot MsgText(), True

  End If
End If
End Function
Public Function OKFinalCust(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, TotalBalance As Double
  Dim M1 As String, M2 As String
  Dim UBCustRecLen As Integer, UBCustF As Integer
  If RecNo& > 0 Then
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCustF = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
  Get UBCustF, RecNo&, UBCustRec(1)
  Close UBCustF

  If UBCustRec(1).Status <> "A" Then
    UBLog "NOFinal:" + Str$(RecNo&) + " NOT ACTIVE"
    M1$ = "This account IS NOT ACTIVE"
    M2$ = "CAN NOT SET THIS ACCOUNT TO FINAL!"
    OKFinalCust = False
  Else
    OKFinalCust = True
  End If
  If OKFinalCust = False Then
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(3).FontSize
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR!"
    MsgText(1) = ""
    MsgText(2) = M1$
    MsgText(3) = ""
    MsgText(4) = M2$
    MsgText(5) = ""
    GetOKorNot MsgText(), True

  End If
End If
End Function

'!!! populates the combo box with revenues
Public Function FillRevList(x As fpCombo)
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  Dim cnt As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  x.AddItem "All Revenues"
  For cnt = 1 To 15
  If Trim(UBSetUpRec(1).Revenues(cnt).RevName) = "" Then
    Exit For
  End If
  x.AddItem Trim(UBSetUpRec(1).Revenues(cnt).RevName)
  Next
  Erase UBSetUpRec
End Function
'!!! from gl common for date check on report screens
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

Public Function GetNumRateRecs%()
  Dim UBRateTblRecLen As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))
  GetNumRateRecs = FileSize(UBPath + "UBRATE.DAT") \ UBRateTblRecLen
  Erase UBRateTblRec
End Function
Public Function GetNumOfRevs%()
  Dim UBSetupLen As Integer, NumofRevs As Integer, Handle As Integer
  Dim RevCnt As Integer, TempRev As String
  NumofRevs = 15
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUp(1))
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  Handle = FreeFile
'  Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
'  Get #Handle, 1, UBSetUpRec(1)
    LoadUBSetUpFile UBSetUp(), UBSetupLen
'this doesn't work properly if they skip around in revenue setup list
  For RevCnt = 1 To 15
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(RevCnt).RevName)
    If Len(TempRev$) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next
  GetNumOfRevs = NumofRevs
  Erase UBSetUp
End Function

Public Sub LoadUBSetUpFile(UBSetUpRec() As UBSetupRecType, UBSetupLen)
  Dim Handle As Integer
  UBSetupLen = Len(UBSetUpRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "UBSETUP.DAT") Then
    Handle = FreeFile
    Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, UBSetUpRec(1)
    End If
    Close Handle
  End If
End Sub
Public Sub LoadCMSetUpFile(CMSetUpRec() As CMSetupType, CMSetuplen)
  Dim Handle As Integer
  CMSetuplen = Len(CMSetUpRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "CMSetTown.dat") Then
    Handle = FreeFile
    Open UBPath$ + "CMSetTown.dat" For Random Shared As Handle Len = CMSetuplen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, CMSetUpRec(1)
    End If
    Close Handle
  End If
End Sub
Public Static Function FillAcctNumName(txtField As fpCombo)
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
  OpenAcctFile AcctFile
  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
  NumAccts = LOF(AcctFile) / Len(GLAcct)
  txtField.Row = -1
  txtField.InsertRow = Str$(0) & Chr$(9) & "Not Found" & Chr$(9) & "Invalid Account" & Chr$(9) & "0"
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
Public Function Exist(FileName$)
    ''' REDIRECTED THIS TO FileExists() function. May need GLOBAL REPLACE to remove all references to this. -sng
''    Exist = FileExists(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  Dim testFile As String
  testFile = UCase$(FileName$)

  On Local Error GoTo FileError

  FileHandle = FreeFile
  Open testFile For Input Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
  End If
  GoTo ExistExit
FileError:
  Close FileHandle
  Exist = False
  If UCase(Error) <> "FILE NOT FOUND" Then
    MsgBox "Error " & Error$ & " " & testFile, vbOKOnly, "Error"
  End If
ExistExit:
  On Error GoTo 0
End Function

'Public Function Exist(FileName$)
'  Dim FileHandle As Integer
'  Dim FileSize As Long
'  FileHandle = FreeFile
'  Open FileName$ For Binary As FileHandle
'  FileSize = LOF(FileHandle)
'  Close FileHandle
'  If FileSize > 0 Then
'    Exist = True
'  Else
'    Exist = False
'    Kill FileName$
'  End If
'End Function

Public Sub KillFile(FileName$)
  If Exist(FileName) Then
    Kill FileName$
  End If
End Sub

Public Function RemNulls$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  RemNulls$ = Text
End Function

Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

Public Static Function Using$(ByVal fmt As String, ByVal Number As Double, Optional LeadZeroFlag As Boolean)
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
  FmtNumber = Space$(Len(fmt))
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen = 0 Then
    TempNumber = "0"
    GoTo GotZero
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
GotZero:
  If LeadZeroFlag Then
    If TempNumber = ".00" Then
      TempNumber = "0.00"
    End If
  End If
  
  RSet FmtNumber = TempNumber
  
  Using = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function

Public Sub MakeSequenceIndex(IndexText$, Parent As Form)
  'Parent.Enabled = False
  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  ReDim UBCustRec(1) As NewUBCustRecType
  
  Dim CustRecLen As Integer, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Integer, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long
  
  CustRecLen = Len(UBCustRec(1))
  
  NumCustRecs& = GetNumOfCust&
  
  ReDim SequenceIndex(1 To NumCustRecs&) As UBSequenceIndexType
  IndexRecLen = Len(SequenceIndex(1))
  
  CHandle = FreeFile
  Open UBCustFile For Random Shared As CHandle Len = CustRecLen
  For cnt = 1 To NumCustRecs&
    Get CHandle, cnt, UBCustRec(1)
    SequenceIndex(cnt).SeqNumber = UBCustRec(1).Seq
    SequenceIndex(cnt).RecNum = cnt
    FrmShowPctComp.ShowPctComp cnt, NumCustRecs&
  Next
  Close CHandle
  
  Load frmInfo
  frmInfo.Label1 = "Sorting. . ."
  DoEvents
  frmInfo.Show
  DoEvents
  SeqQSort SequenceIndex(), 1, NumCustRecs&
  Unload frmInfo
  DoEvents
  
  FrmShowPctComp.Label1 = "Writing Customer Index."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show

  KillFile TempIndexName
  IHandle = FreeFile
  Open TempIndexName For Random Shared As IHandle Len = 4
  
  For cnt = 1 To NumCustRecs&
    Prec& = SequenceIndex(cnt).RecNum
    Put IHandle, cnt, Prec&
    FrmShowPctComp.ShowPctComp cnt, NumCustRecs&
  Next
  Close IHandle
  
  Erase UBCustRec, SequenceIndex
  'Parent.Enabled = True
End Sub
Public Sub MakeMowZipCodeIndex(IndexText$)
  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim CustRecLen As Integer, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Integer, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long, NumOfBillRec As Long
  Dim BCnt As Long
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
  For BCnt = 1 To NumOfBillRec
    Get CHandle, BCnt, UBCustRec(1)
    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
    ZipIndex(BCnt).RecNum = BCnt
    FrmShowPctComp.ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
  Next
  Close
  Load frmInfo
  frmInfo.Label1 = "Sorting. . ."
  DoEvents
  frmInfo.Show
  DoEvents
  ZipQSort ZipIndex(), 1, NumOfBillRec
  Unload frmInfo
  DoEvents
  
  FrmShowPctComp.Label1 = "Writing Index Records."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show

 KillFile TempIndexName
  IHandle = FreeFile
  Open TempIndexName For Output As IHandle
  Close IHandle

  IHandle = FreeFile
  Open TempIndexName For Random Shared As IHandle Len = 4
  For cnt = 1 To NumOfBillRec
    Prec& = ZipIndex(cnt).RecNum
    Put IHandle, cnt, Prec&
    FrmShowPctComp.ShowPctComp cnt, NumOfBillRec               'show user percentage complete
  Next
  Close IHandle

  Erase UBCustRec, ZipIndex

End Sub
Public Sub MakeZipCodeIndex(IndexText$)
'Removed all rest of code
  Call MakeMowZipCodeIndex(IndexText$)

End Sub
    
'For Mail Lables
Public Sub MakePostalIndex(IndexText$)
  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  
  Dim CustRecLen As Integer, IndexRecLen As Integer
  Dim CHandle As Integer, cnt As Integer, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long
  Dim BCnt As Long

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumCustRecs = GetNumOfCust

  ReDim PostalIndex(1 To NumCustRecs) As UBPostalIndexType
  IndexRecLen = Len(PostalIndex(1))

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For cnt = 1 To NumCustRecs
    Get CHandle, cnt, UBCustRec(1)
    PostalIndex(cnt).ZIPCODE = UBCustRec(1).ZIPCODE
    RSet PostalIndex(cnt).Route = QPTrim$(UBCustRec(1).POSTRTE)
    PostalIndex(cnt).RecNum = cnt
    FrmShowPctComp.ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next

  Close CHandle
  Load frmInfo
  frmInfo.Label1 = "Sorting. . ."
  DoEvents
  frmInfo.Show
  DoEvents
  PostalQSort PostalIndex(), 1, NumCustRecs
  Unload frmInfo
  DoEvents
  
  FrmShowPctComp.Label1 = "Writing Index Records."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show

  IHandle = FreeFile

  'FCreate TempIndexName
  KillFile TempIndexName
  Open TempIndexName For Random Shared As IHandle Len = 4
  For cnt = 1 To NumCustRecs
    Prec& = PostalIndex(cnt).RecNum
    Put IHandle, cnt, Prec&
    FrmShowPctComp.ShowPctComp cnt, NumCustRecs                'show user percentage complete
  Next
  Close IHandle

  Erase UBCustRec, PostalIndex
End Sub
'Function returns True if a customer has been deleted.
Public Function IsDeleted%(AcctNum&)
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim Handle As Integer
  Dim UBCustRecLen As Integer
  
  UBCustRecLen = Len(UBCustRec(1))
  Handle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As Handle Len = UBCustRecLen
  Get Handle, AcctNum&, UBCustRec(1)
  Close Handle
  
  If UBCustRec(1).DelFlag <> 0 Then
    IsDeleted% = True
  Else
    IsDeleted% = False
  End If
  Erase UBCustRec

End Function
Public Function IsDCDeleted%(AcctNum&)
  ReDim DCCustREc(1) As DCCustRecType
  Dim Handle As Integer
  Dim DCCustRecLen As Integer
  
  DCCustRecLen = Len(DCCustREc(1))
  Handle = FreeFile
  Open "DCCUST.DAT" For Random Shared As Handle Len = DCCustRecLen
  Get Handle, AcctNum&, DCCustREc(1)
  Close Handle
  
  If UCase$(DCCustREc(1).Deleted) <> "Y" Then
    IsDCDeleted% = False
  Else
    IsDCDeleted% = True
  End If
  Erase DCCustREc

End Function

'This function returns the number of customer records
Public Function GetNumOfCust&()
  ReDim TCustRec(1) As NewUBCustRecType
  Dim RecLen As Integer
  RecLen = Len(TCustRec(1))
  GetNumOfCust = FileSize(UBCustFile) \ RecLen
  Erase TCustRec
End Function
  
Public Function FileSize(FileName$) As Long
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
  End If
End Function

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
  frmLoadingRpt.Show
  DoEvents
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
  DoEvents
  Unload frmLoadingRpt
  DoEvents
  frmViewPrint.Show vbModal
End Sub

Public Function GetDefaultLookUP%()
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  GetDefaultLookUP = Val(UBSetUpRec(1).DefLook)
  Erase UBSetUpRec
End Function

'Function to format the Book part of a location number
Public Function FmtBook$(BOOK$)
  Dim BookLen As Integer
  
  BOOK$ = QPTrim$(BOOK$)
  BookLen = Len(BOOK$)
  
  Select Case BookLen
  Case 0
    FmtBook$ = "00"
  Case 1
    FmtBook$ = "0" + BOOK$
  Case Else
    FmtBook$ = BOOK$
  End Select
  
End Function

'Function to format the Sequence part of a location number
Public Function FmtSeqN$(SeqN$)
  Dim TSeq As String
  Dim SeqNLen As Integer
  
  SeqN$ = QPTrim$(SeqN$)
  SeqNLen = Len(SeqN$)
  
  Select Case SeqNLen
  Case 0
    FmtSeqN$ = "000000"
  Case 1 To 5
    TSeq = "000000" + SeqN$
    FmtSeqN$ = Right$(TSeq$, 6)
  Case Else
    FmtSeqN$ = SeqN$
  End Select
End Function

Public Function GetCustMeterType(UBCustRec() As NewUBCustRecType, ThisMeter) As Integer
  
  Dim LMtrType    As String
  Dim LMtrTypeLen As Integer, LThisMeter As Integer
  
  'Meter Types
  'CONST MtrWaterOnly = 1
  'CONST MtrSewerOnly = 2
  'CONST MtrCombined = 3
  'CONST MtrElectric = 4
  'CONST MtrDemand = 5
  'CONST MtrGas = 6
  'CONST MtrTouchRead = 7
  
  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MtrType)
  LMtrTypeLen = Len(LMtrType$)
  If LMtrTypeLen > 0 Then
    Select Case LMtrType$
    Case "W"
      LThisMeter = MtrWaterOnly
    Case "S"
      LThisMeter = MtrSewerOnly
    Case "C"
      LThisMeter = MtrCombined
    Case "E"
      LThisMeter = MtrElectric
    Case "D"
      LThisMeter = MtrDemand
    Case "G"
      LThisMeter = MtrGas
    Case "T"
      LThisMeter = MtrTouchRead
    Case Else
      LThisMeter = True
    End Select
    GetCustMeterType = LThisMeter
  Else
    GetCustMeterType = 0
  End If
  
End Function

Public Sub CMTerminate()
  Dim CMFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  On Local Error Resume Next
  CMLog "CM Exited: "
  Ready4others PWcnt
  If DebugMode = False Then
    Shell "CitiPak.exe", vbMaximizedFocus
  End If
  For CMFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(CMFrmCnt)
  Next
  DoTheTime
  DoEvents
  End
End Sub
Public Sub CitiTerminate()
  Dim CMFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  ClearInUse PWcnt
  For CMFrmCnt = Forms.Count - 1 To 0 Step -1
    Unload Forms(CMFrmCnt)
  Next
  DoEvents
  End
End Sub
Public Static Sub UBLog(Text$)
  Dim Today As String, TheTime As String
  Dim AmPm As String, Hour As String
  Dim ThisHour As Integer, LogFile As Integer
  
  Today$ = Date$
  Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    ThisHour = Val(Mid$(TheTime$, 2, 1))
  Else
    ThisHour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case ThisHour
  Case Is > 11
    ThisHour = ThisHour - 12
    If ThisHour = 0 Then ThisHour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case ThisHour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(Str$(ThisHour))
    Case Else
      Hour$ = QPTrim$(Str$(ThisHour))
  End Select
  TheTime$ = Hour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open UBPath$ + "UBLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "CM: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub

Public Static Sub CMLog(Text$)
  Dim Today As String, TheTime As String
  Dim AmPm As String, Hour As String
  Dim ThisHour As Integer, LogFile As Integer
  
  Today$ = Date$
  Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    ThisHour = Val(Mid$(TheTime$, 2, 1))
  Else
    ThisHour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case ThisHour
  Case Is > 11
    ThisHour = ThisHour - 12
    If ThisHour = 0 Then ThisHour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case ThisHour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(Str$(ThisHour))
    Case Else
      Hour$ = QPTrim$(Str$(ThisHour))
  End Select
  TheTime$ = Hour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open UBPath$ + "CMLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "CM: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub
Public Sub BLLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "arlog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " CM "; Info$; ; " "; AcctLogFileName
  Close AcctLogFile
End Sub

Public Sub DisplayCustTransList(CustRec As Long)
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim PrevTranRec As Long
  Dim UBFile As Integer, dcnt As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double, PreBal As Double
  
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  UBFile = FreeFile
  Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile

  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
'
Top:
'
  UBFile = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      dcnt = dcnt + 1
      Get UBFile, PrevTranRec&, UBTranRec(1)
      LSet Build = " " + Num2Date(UBTranRec(1).TransDate)
      GoSub GetTransType
      Mid$(Build, 20) = TType$
      Mid$(Build, 48) = Using("#####.##", UBTranRec(1).TransAmt, True)
'      'this will show th actual trans number in the list
'      'MID$(MChoice(DCnt).V, 50) = FUsing(STR$(PrevTranRec&), "######")
'      Mid$(Build, 55) = Str$(PrevTranRec&)
      Mid$(Build, 63) = Using("#####.##", UBTranRec(1).RunBalance, True)
      Mid$(Build$, 71) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      frmTRDispList.fpTRList.AddItem Build$
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop
  End If
  Close UBFile
  frmTRDispList.Label5.Caption = QPTrim(UBCustRec(1).CustName)
  frmTRDispList.Label2 = "Balance: " + Using("#####.##", CurBal# + PreBal#, True)
  frmTRDispList.Label3 = "Current:  " + Using("#####.##", CurBal#, True)
  frmTRDispList.Label4 = "Previous:  " + Using("#####.##", PreBal#, True)
  Unload frmInfo
  DoEvents
  frmTRDispList.Show vbModal
  Erase UBTranRec, UBCustRec

Exit Sub

GetTransType:
'
  Select Case UBTranRec(1).TransType
  Case TranUtilityBill, TranUtilityBill + 100
    TType$ = "Utility Bill "
  Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
    TType$ = "Penalty, Reconnect Fee"
  Case TranBillPayment, TranBillPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "PAYMENT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Utility Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Utility Payment"
    End If
'  Case TranPenaltyPayment
'    TType$ = "Penalty Payment"
  Case TranPenaltyCharge
    TType$ = "Penalty/Late Fee"
  Case TranAppliedDeposit
    TType$ = "Applied Deposit"
  Case TranDepositPayment, TranDepositPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Deposit Payment"
    End If
  Case TranDraftPayment
    TType$ = "Draft Payment"
  Case TranBeginBalance, TranBeginBalance + 100
    TType$ = "Beginning Balance"
  Case 9
    TType$ = "Deposit Refund"
  Case TranUpwardAdjustment
    TType$ = "Upward Adjustment"
  Case TranDownwardAdjustment
    TType$ = "Downward Adjustment"
  Case TranOverPayAdjustment
    TType$ = "Payment Adjustment"
  Case TranDepCreditRemoval
    TType$ = "DepCrRemvl " + Left$(QPTrim$(UBTranRec(1).BillMsg), 10)
  Case TranDepPaymentVoid
    TType$ = "DepPayVoid " + Left$(QPTrim$(UBTranRec(1).BillMsg), 10)
  Case Else
    TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select

Return

End Sub
Public Sub DisplayDCCustTransList(CustRec As Long)
  ReDim DCTranRec(1) As DCTransRecType
  ReDim DCCustREc(1) As DCCustRecType
  Dim DCCustRecLen As Integer, DCTranRecLen As Integer
  Dim PrevTranRec As Long
  Dim DCFile As Integer, dcnt As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double
  
  frmInfo.Label1 = "Loading. . ."
  frmInfo.Show
  DoEvents

  DCCustRecLen = Len(DCCustREc(1))
  DCTranRecLen = Len(DCTranRec(1))
  
  DCFile = FreeFile
  Open UBPath + "DCCust.dat" For Random Shared As DCFile Len = DCCustRecLen
  Get DCFile, CustRec&, DCCustREc(1)
  Close DCFile

  CurBal# = DCCustREc(1).AcctBal
'
Top:
'
  DCFile = FreeFile
  Open UBPath + "DCTRANS.DAT" For Random Shared As DCFile Len = DCTranRecLen
  
  PrevTranRec& = DCCustREc(1).FirstTrans
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      dcnt = dcnt + 1
      Get DCFile, PrevTranRec&, DCTranRec(1)
      LSet Build = Str(DCTranRec(1).TransDate) + Chr9$ + " " + Num2Date(DCTranRec(1).TransDate)
      GoSub GetTransType
      Mid$(Build, 22) = TType$
      Mid$(Build, 53) = Using("#####.##", DCTranRec(1).TransAmount, True)
'      'this will show the actual trans number in the list
'      Mid$(Build, 55) = Str$(PrevTranRec&)
      Mid$(Build, 65) = Using("#####.##", DCTranRec(1).BalanceAfterTrans, True)
      Mid$(Build$, 73) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      frmTRDispListDC.fpTRList.AddItem Build$
      PrevTranRec& = DCTranRec(1).NextTrans
    Loop
  End If
  Close DCFile
  frmTRDispListDC.Label5.Caption = QPTrim(DCCustREc(1).BILLNAME)
  frmTRDispListDC.Label7 = "Acct: " & Str$(CustRec&)
  frmTRDispListDC.Label2 = "Balance: " + Using("#####.##", CurBal#, True)
  'frmTRDispListDC.Label3 = "Current: " + Using("#####.##", CurBal#, True)
  'frmTRDispListDC.Label4 = "Previous: " + Using("#####.##", PreBal#, True)
  Unload frmInfo
  DoEvents
  frmTRDispListDC.Show vbModal
  Erase DCTranRec, DCCustREc

Exit Sub

GetTransType:
'
  Select Case DCTranRec(1).TransType
  Case 1 'Charge
    TType$ = "Decal Charge"
  Case 2 'Payment
    TType$ = "Decal Payment"
  Case 3  'Charge Void
    TType$ = "Void Charge"
  Case 4  'Payment Void
    TType$ = "Void Payment"
  Case Else
    TType$ = Str$(DCTranRec(1).TransType) + " ???"
  End Select
  TDesc$ = QPTrim$(DCTranRec(1).TRVinDesc)
Return

End Sub
Public Sub PrintTRListScreenDC()
  Unload frmTRDispListDC
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "DCTRlist.rpt", "Customer Transaction List"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "DCTRlist.rpt"
    ARptLineRpt.startrpt
  End If
End Sub
Public Sub PrintTRDetlScreenDC()
  Unload frmTRDetailDC
  Unload frmTRDispListDC
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "DCTRDetl.RPT", "Customer Detail Transaction"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "DCTRDetl.RPT"
    ARptLineRpt.startrpt
  End If
End Sub

Public Function CustHasMsg(RecNo&)
  
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim MsgRec(1) As UBMessRecType
  Dim MsgLen As Integer, UBCustRecLen As Integer
  Dim UBFile As Integer, zz As Integer
  Dim NumMsgRec As Long, MRec As Long
  
  CustHasMsg = False
  
  MsgLen = Len(MsgRec(1))
  NumMsgRec& = FileSize&("UBMESAGE.DAT") / MsgLen

  UBCustRecLen = Len(UBCustRec(1))

  If RecNo& > 0 Then
    UBFile = FreeFile
    Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
    Get UBFile, RecNo&, UBCustRec(1)
    Close UBFile
    MRec& = UBCustRec(1).MessageRec
    If MRec& > 0 And MRec& <= NumMsgRec& Then
      UBFile = FreeFile
      Open UBPath + "UBMESAGE.DAT" For Random Shared As UBFile Len = MsgLen
      Get UBFile, MRec&, MsgRec(1)
      Close UBFile
      For zz = 1 To 15
        'QPTrim$ (MsgRec(1).MessLine(zz).Line)
        If Len(QPTrim$(MsgRec(1).MessLine(zz).Msg)) > 0 Then
          CustHasMsg = True
          Exit For
        End If
      Next
    End If
  End If
  
  Erase UBCustRec, MsgRec
  
End Function

Public Function GetOKorNot%(MsgText() As String, Optional OKOnly As Boolean)
  Dim zz As Integer, RetValue As Integer
  If OKOnly Then
    frmMsgDialog.RetLabel = "-2"
  End If
  frmMsgDialog.Caption = MsgText(0)
  For zz = 1 To 5
    frmMsgDialog.Label(zz - 1) = MsgText(zz)
  Next
  frmMsgDialog.Show vbModal
  RetValue = Val(frmMsgDialog.RetLabel)
  Unload frmMsgDialog
  GetOKorNot% = RetValue
End Function
'
Public Function ErrorScrn(WhatError%, Acct&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer

  ErrorScrn = True

  Select Case WhatError
  Case 1
    MsgText(3) = "Has Invalid Reading!"
  Case 2
    MsgText(3) = "Invalid Book Number!"
  Case 3
    MsgText(3) = "Has an INVALID RATE CODE!!"
  Case 4
    MsgText(3) = "Has Mismatched Meters!"
  Case 5
    MsgText(3) = "Has an INVALID Reading!"
  Case 6
    MsgText(3) = "INVALID Flat Rate Info!"
  Case 7
    MsgText(3) = "INVALID Monthly Billed Code!"
  Case 8
    MsgText(3) = "Meters with NO RATE Code!"
  Case 9
    MsgText(3) = "Invalid Customer Type!"
  End Select
  MsgText(0) = "ERROR:"
  MsgText(1) = "Account Number: " + Str$(Acct&)
  MsgText(2) = ""
  MsgText(4) = ""
  MsgText(5) = "Correct and Try Again."
  GetOKorNot MsgText(), True

 ' QPrintRC "ACCOUNT:" + Str$(Acct&), 10, AcCol, -1
 ' QPrintRC "Correct and Print Again.", 13, 28, -1

 ' ShowCursor
 ' Get.Moose.OR.Key Ky$, MooseButton%, MRow%, MCol%

'  If Len(Ky$) = 2 Then
'    If Right$(Ky$, 1) = "g" Then
    
 '     ErrorScrn = False
      'LPRINT Acct&
 '   End If
'  End If
'  RestScrn TempArray()
'  Erase TempArray
'this code below came from custaddedit form
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(2).FontSize
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = ""
'    MsgText(3) = "There are NO transactions to display."
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True

End Function

Public Sub OpenMiscCodeFile(NumOfMiscRecs)
  Dim MCFile As Integer, MiscCodeRecLen As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen

End Sub
Public Sub OpenAcctFile(AcctFileNum, Optional NumAccts As Integer)
  Dim GLAcctRecLen As Integer
  GLAcctRecLen = Len(GLAcct)
  AcctFileNum = FreeFile
  Open UBPath$ + "GLAcct.DAT" For Random Shared As AcctFileNum Len = GLAcctRecLen
  NumAccts = LOF(AcctFileNum) \ GLAcctRecLen
End Sub
 Public Sub PrintCustInfo(Rec As Long, RptType As Integer)
  Dim PageNo As Integer, Title As String, tb As Integer
  Dim Dash80 As String, ReportFile As String
  Dim UBRpt As Integer, ToPrint As String
  Dim Msgflag As Boolean, RecNo As Long, NumOfRates As Integer, cnt As Integer
  Dim tmpCustRec As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  Title$ = "Customer Information Report"
  Dash80$ = String$(80, "-")

  ReportFile$ = UBPath$ + "UBINFRPT.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  CustRecLen = Len(tmpCustRec)
  Dim UBSetupLen As Integer
  Dim RevCnt As Integer, GCode As String
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUpRec(1))
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer
  RecNo& = Rec
  UBHandle = FreeFile
  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen

  Get #UBHandle, RecNo&, tmpCustRec
  Close UBHandle
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  NumOfRates = GetNumRateRecs%
  GrpCodeRecLen = Len(GroupCde)

  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  If tmpCustRec.GroupCodeRec > 0 Then
    Get #ghandle, tmpCustRec.GroupCodeRec, GroupCde
    If GroupCde.Deleted = 0 Then
      GCode$ = QPTrim$(GroupCde.GroupCODE) + " " + QPTrim$(GroupCde.GroupCodeName)
    Else
      GCode$ = QPTrim$(GroupCde.GroupCODE) + " Inactive"
    End If
  Else
    GCode$ = "None"
  End If
  Close #ghandle

  If CustHasMsg(RecNo) Then
    Msgflag = True
    'MsgRec = tmpCustRec.MessageRec
  End If

  If RptType = 1 Then 'do the graphics
  ToPrint$ = ""
  ToPrint$ = Str$(RecNo) + "~" + tmpCustRec.BOOK + "~" + tmpCustRec.SEQNUMB
  ToPrint$ = ToPrint$ + "~" + tmpCustRec.Status + "~" + Num2Date(tmpCustRec.OPENDATE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.SEARCH)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.CustName)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.Addr1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.Addr2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.SERVADDR)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.City)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.State)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.ZIPCODE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.DPCode)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HPHONE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.WPHONE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.SOSEC)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.DRVLIC)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.CUSTTYPE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.Addr911)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.BillTo)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.BILLCOPY))
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.POSTRTE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.BILLCYCL))
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.ZONE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Seq))
  Select Case tmpCustRec.CASHONLY
  Case "N", " "
    ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
    ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  Select Case tmpCustRec.LATEFEE
  Case "N", " "
    ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
    ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select

  Select Case tmpCustRec.CUTOFFYN
  Case "N", " "
    ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
    ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  Select Case tmpCustRec.TAXEXPT
  Case "N", " "
     ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
     ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.SRCIT)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.USEDRAFT)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.AcctType)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.BankName)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.BANKLOC)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.TRANSIT)
  ToPrint$ = ToPrint$ + "~" + "XXXXXXXXXXXX"
  ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", Str$((Round#(tmpCustRec.CurrBalance + tmpCustRec.PrevBalance))))
  ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", Str$(Round#(tmpCustRec.CurrBalance)))
  ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", Str$(Round#(tmpCustRec.PrevBalance)))
  ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", Str$(Round#(tmpCustRec.DepositAmt)))

  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.BILLCMNT)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.PAYCMNT)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.PumpCode)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.USERCODE1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.USERCODE2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.ProratePCT))
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG3)

  For cnt = 0 To 14
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBSetUpRec(1).Revenues(cnt + 1).RevName)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.serv(cnt + 1).RATECODE)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.serv(cnt + 1).RMtrType)
  Next
  For cnt = 0 To 3
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRDESC)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).FRAMT))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).REVSRC))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).NumMin))
  Next
  For cnt = 0 To 1
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).AmtOwed))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).TotAmtPD))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).PayAmt))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).RevSource))
  Next
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.MFEE1))
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.MFEE2))

  For cnt = 0 To 6
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrNum)
    If tmpCustRec.LocMeters(cnt + 1).MTRMulti > 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).MTRMulti))
    Else
      ToPrint$ = ToPrint$ + "~" + " "
    End If
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrType)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.LocMeters(cnt + 1).MTRUnit)
    If tmpCustRec.LocMeters(cnt + 1).NumUser > 0 Then
      ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).NumUser))
    Else
      ToPrint$ = ToPrint$ + "~" + " "
    End If
    ToPrint$ = ToPrint$ + "~" + Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate)
    If tmpCustRec.LocMeters(cnt + 1).CurRead > 0 Then
      ToPrint$ = ToPrint$ + "~" + Str$(tmpCustRec.LocMeters(cnt + 1).CurRead)
    Else
     ToPrint$ = ToPrint$ + "~" + " "
    End If
    If tmpCustRec.LocMeters(cnt + 1).PrevRead > 0 Then
      ToPrint$ = ToPrint$ + "~" + Str$(tmpCustRec.LocMeters(cnt + 1).PrevRead)
    Else
     ToPrint$ = ToPrint$ + "~" + " "
    End If
    ToPrint$ = ToPrint$ + "~" + Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate)
    ToPrint$ = ToPrint$ + "~" + Num2Date(tmpCustRec.LocMeters(cnt + 1).PastDate)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrIDNO)
'put new field here
    ToPrint$ = ToPrint$ + "~" + Str$(tmpCustRec.LocMeters(cnt + 1).MtrLat)
    ToPrint$ = ToPrint$ + "~" + Str$(tmpCustRec.LocMeters(cnt + 1).MtrLng)
  Next
    ToPrint$ = ToPrint$ + "~" + GCode$
  Print #UBRpt, ToPrint$
  Close
  Load frmLoadingRpt
  'frmLoadingRpt.setwherefrom frmUBCustMenu
  ARptUBCustInfo.txtDate = Now
  ARptUBCustInfo.txtTown = TownName$
  ARptUBCustInfo.GetName ReportFile$
  ARptUBCustInfo.startrpt
  Else
  Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
  Print #UBRpt, Tab(30); Title$
  Print #UBRpt, Now
  Print #UBRpt, TownName$
  Print #UBRpt, Dash80$
  Print #UBRpt,
  Print #UBRpt, "Customer Name: "; QPTrim$(tmpCustRec.CustName);
  Print #UBRpt, Tab(46); "Status: "; tmpCustRec.Status
  Print #UBRpt, "Account #: "; Str$(RecNo); Tab(25); "Location: "; tmpCustRec.BOOK; "-"; tmpCustRec.SEQNUMB;
  Print #UBRpt, Tab(50); "Account Opened: "; Num2Date(tmpCustRec.OPENDATE)
  Print #UBRpt, "Address: "; QPTrim$(tmpCustRec.Addr1);
  Print #UBRpt, Tab(50); "Group Code:  " & GCode$
  Print #UBRpt, Tab(9); QPTrim$(tmpCustRec.Addr2);
  Print #UBRpt, Tab(46); "----Account Balance Information----"
  Print #UBRpt, Tab(9); QPTrim$(tmpCustRec.City); " "; QPTrim$(tmpCustRec.State); " "; QPTrim$(tmpCustRec.ZIPCODE);
  Print #UBRpt, Tab(50); "Account Balance: "; Using$("$###,###,###.##", Str$((Round#(tmpCustRec.CurrBalance + tmpCustRec.PrevBalance))))
  Print #UBRpt, "Service Address: "; QPTrim$(tmpCustRec.SERVADDR);
  Print #UBRpt, Tab(50); "       Past Due: "; Using$("$###,###,###.##", Str$(Round#(tmpCustRec.PrevBalance)))
  Print #UBRpt, Tab(50); "        Current: "; Using$("$###,###,###.##", Str$(Round#(tmpCustRec.CurrBalance)))
  Print #UBRpt, "Home Phone: "; QPTrim$(tmpCustRec.HPHONE);
  Print #UBRpt, Tab(50); " Amt on Deposit: "; Using$("$###,###,###.##", Str$(Round#(tmpCustRec.DepositAmt)))
  Print #UBRpt, "Work Phone: "; QPTrim$(tmpCustRec.WPHONE);
  Print #UBRpt, Tab(46); "-------- Draft Information -------"
  Print #UBRpt, "Search Name: "; QPTrim$(tmpCustRec.SEARCH);
  Print #UBRpt, Tab(50); "      Use Draft: "; QPTrim$(tmpCustRec.USEDRAFT)
  Print #UBRpt, "DPCode: "; QPTrim$(tmpCustRec.DPCode);
  Print #UBRpt, Tab(50); "  Draft Account: "; QPTrim$(tmpCustRec.AcctType)
  Print #UBRpt, "SocSecNo: "; QPTrim$(tmpCustRec.SOSEC);
  Print #UBRpt, Tab(50); "      Bank Name: "; QPTrim$(tmpCustRec.BankName)
  Print #UBRpt, "Driver Lic#: "; QPTrim$(tmpCustRec.DRVLIC);
  Print #UBRpt, Tab(50); "  Bank Location: "; QPTrim$(tmpCustRec.BANKLOC)
  Print #UBRpt, "Customer Type: "; QPTrim$(tmpCustRec.CUSTTYPE);
  Print #UBRpt, Tab(50); "        Transit: "; QPTrim$(tmpCustRec.TRANSIT)
  Print #UBRpt, "911 Addr: "; QPTrim$(tmpCustRec.Addr911);
  Print #UBRpt, Tab(50); "   Bank Account: "; "XXXXXXXXXXXX"

  Print #UBRpt, "Bill To: "; QPTrim$(tmpCustRec.BillTo)
  Print #UBRpt, "Bill Copies: "; QPTrim$(Str$(tmpCustRec.BILLCOPY));
  Print #UBRpt, Tab(39); "---------- Service Information ---------"
  Print #UBRpt, "Postal Route: "; QPTrim$(tmpCustRec.POSTRTE);
  Print #UBRpt, Tab(39); " Rev                 Rate         MtrType"
  Print #UBRpt, "Bill Cycle: "; QPTrim$(Str$(tmpCustRec.BILLCYCL));
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(1).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(1).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(1).RMtrType)
  Print #UBRpt, "Zone: "; QPTrim$(tmpCustRec.ZONE);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(2).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(2).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(2).RMtrType)
  Print #UBRpt, "Read Seq: "; QPTrim$(Str$(tmpCustRec.Seq));
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(3).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(3).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(3).RMtrType)

  Select Case tmpCustRec.CASHONLY
  Case "N", " "
    Print #UBRpt, "Cash Only: "; "No";
  Case Else
    Print #UBRpt, "Cash Only: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(4).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(4).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(4).RMtrType)
  Select Case tmpCustRec.LATEFEE
  Case "N", " "
    Print #UBRpt, "Late Fee: "; "No";
  Case Else
    Print #UBRpt, "Late Fee: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(5).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(5).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(5).RMtrType)

  Select Case tmpCustRec.CUTOFFYN
  Case "N", " "
    Print #UBRpt, "Allow Cutoff: "; "No";
  Case Else
    Print #UBRpt, "Allow Cutoff: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(6).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(6).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(6).RMtrType)
  Select Case tmpCustRec.TAXEXPT
  Case "N", " "
    Print #UBRpt, "Tax Exempt: "; "No";
  Case Else
    Print #UBRpt, "Tax Exempt: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(7).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(7).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(7).RMtrType)

  Print #UBRpt, "Senior Citizen: "; QPTrim$(tmpCustRec.SRCIT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(8).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(8).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(8).RMtrType)

  Print #UBRpt, "Bill Comment: "; QPTrim$(tmpCustRec.BILLCMNT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(9).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(9).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(9).RMtrType)

  Print #UBRpt, "Pay Comment: "; QPTrim$(tmpCustRec.PAYCMNT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(10).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(10).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(10).RMtrType)

  Print #UBRpt, "Pump Code: "; QPTrim$(tmpCustRec.PumpCode);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(11).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(11).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(11).RMtrType)

  Print #UBRpt, "User Code 1: "; QPTrim$(tmpCustRec.USERCODE1);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(12).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(12).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(12).RMtrType)

  Print #UBRpt, "User Code 2: "; QPTrim$(tmpCustRec.USERCODE2);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(13).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(13).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(13).RMtrType)

  Print #UBRpt, "Prorate%: "; QPTrim$(Str$(tmpCustRec.ProratePCT));
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(14).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(14).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(14).RMtrType)

  Print #UBRpt, "HH Message 1: "; QPTrim$(tmpCustRec.HHMSG1);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(15).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(15).RATECODE);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(15).RMtrType)

  Print #UBRpt, "HH Message 2: "; QPTrim$(tmpCustRec.HHMSG2);
  Print #UBRpt, Tab(45); "MembFee Refundable - "; QPTrim$(Str$(tmpCustRec.MFEE1))
  Print #UBRpt, "HH Message 3: "; QPTrim$(tmpCustRec.HHMSG3);
  Print #UBRpt, Tab(45); "MembFee NonRef - "; QPTrim$(Str$(tmpCustRec.MFEE2))

  Print #UBRpt, "-------- Flat Rate  Information -------";
  Print #UBRpt, Tab(45); "--------- Monthly Payments --------"
  Print #UBRpt, "Desc        Amt       Freq     Rev  Min";
  Print #UBRpt, Tab(45); "Amt Owed   Amt Paid   Payment  Rev"

  For cnt = 0 To 1
    Print #UBRpt, Left$(tmpCustRec.FlatRates(cnt + 1).FRDESC, 10);
    Print #UBRpt, Tab(13); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).FRAMT));
    Print #UBRpt, Tab(22); QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ);
    Print #UBRpt, Tab(33); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).REVSRC));
    Print #UBRpt, Tab(37); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).NumMin));
    Print #UBRpt, Tab(48); QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).AmtOwed));
    Print #UBRpt, Tab(60); QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).TotAmtPD));
    Print #UBRpt, Tab(70); QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).PayAmt));
    Print #UBRpt, Tab(78); QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).RevSource))

  Next

  For cnt = 2 To 3
    Print #UBRpt, Left$(tmpCustRec.FlatRates(cnt + 1).FRDESC, 10);
    Print #UBRpt, Tab(13); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).FRAMT));
    Print #UBRpt, Tab(22); QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ);
    Print #UBRpt, Tab(33); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).REVSRC));
    Print #UBRpt, Tab(37); QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).NumMin))
  Next
  Print #UBRpt,
  Print #UBRpt, "Meter Information -------------"
'  Print #UBRpt, "Mtr 1"; Tab(14); "Mtr 2"; Tab(26); "Mtr 3"; Tab(37); "Mtr 4";
'  Print #UBRpt, Tab(48); "Mtr 5"; Tab(59); "Mtr 6"; Tab(70); "Mtr 7"

  Print #UBRpt, "   MtrN   Mult T U N  InstDate  CurRead PrvRead  CurrDate   PrevDate   IDNo  Lat   Long"
  For cnt = 0 To 6
    Print #UBRpt, QPTrim$(Str$(cnt + 1)); ")"; Left$(QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrNum), 8);
    Print #UBRpt, Tab(12); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).MTRMulti));
    Print #UBRpt, Tab(16); QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrType);
    Print #UBRpt, Tab(18); QPTrim$(tmpCustRec.LocMeters(cnt + 1).MTRUnit);
    Print #UBRpt, Tab(20); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).NumUser));
    Print #UBRpt, Tab(22); Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate);
    Print #UBRpt, Tab(33); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).CurRead));
    Print #UBRpt, Tab(41); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).PrevRead));
    Print #UBRpt, Tab(49); Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate);
    Print #UBRpt, Tab(60); Num2Date(tmpCustRec.LocMeters(cnt + 1).PastDate);
    Print #UBRpt, Tab(71); QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrIDNO);
    Print #UBRpt, Tab(77); Str$(tmpCustRec.LocMeters(cnt + 1).MtrLat);
    Print #UBRpt, Tab(83); Str$(tmpCustRec.LocMeters(cnt + 1).MtrLng)
  Next

  Print #UBRpt,
  Print #UBRpt, Dash80$
  Print #UBRpt, Chr$(12)

  Close

  ViewPrint ReportFile$, Title$
  KillFile ReportFile$
  End If
End Sub
 Public Sub PrintDCCustInfo(Rec As Long, RptType As Integer)
  Dim PageNo As Integer, Title As String, tb As Integer, dcnt As Long
  Dim Dash80 As String, ReportFile As String, Num1 As Long, LineCnt As Integer
  Dim DCRpt As Integer, ToPrint As String, TPDate As String
  Dim Msgflag As Boolean, RecNo As Long, NumOfVehs As Long, cnt As Long
  Dim NumOfDCRecs As Long, DCFile As Integer, GCode As String
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer
  Dim MaxLine As Integer
  LineCnt = 0
  ReDim DCCustREc(1) As DCCustRecType
  RecNo = Rec
  OpenDCCustFile NumOfDCRecs, DCFile
  Get DCFile, RecNo, DCCustREc(1)
  Close DCFile
  Title$ = "Customer Information Report"
  Dash80$ = String$(80, "-")
  TPDate$ = ""
  ReportFile$ = UBPath + "DCINFo.RPT"
  DCRpt = FreeFile
  Open ReportFile$ For Output As DCRpt
  ToPrint$ = ""
  MaxLines = 60
  If RptType = 1 Then 'do the graphics
  ToPrint$ = ""
  ToPrint$ = Str$(RecNo) + "~" + QPTrim$(DCCustREc(1).CUSTNUMB)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).SORTNAME)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).BILLNAME)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).ADDRESS1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).ADDRESS2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).City)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).State)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).ZIPCODE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).SOSEC)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).DRVLIC)
  ToPrint$ = ToPrint$ + "~" + Num2Date(DCCustREc(1).DATEOPED)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).HPHONE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(DCCustREc(1).WPHONE)
  Select Case DCCustREc(1).CASHONLY
    Case "N", " "
    ToPrint$ = ToPrint$ + "~" + "No"
  Case Else
    ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  If DCCustREc(1).FirstCar > 0 Then
    ToPrint$ = ToPrint$ + "~" + "Yes"
  Else
    ToPrint$ = ToPrint$ + "~" + "No"
  End If
  Select Case DCCustREc(1).resident
    Case "N", " "
      ToPrint$ = ToPrint$ + "~" + "No"
    Case Else
      ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
  Select Case DCCustREc(1).Owner
    Case "N", " "
      ToPrint$ = ToPrint$ + "~" + "No"
    Case Else
      ToPrint$ = ToPrint$ + "~" + "Yes"
  End Select
   ToPrint$ = ToPrint$ + "~" + Using$("$###,###,###.##", DCCustREc(1).AcctBal)
  ReDim DCVRec(1) As DCVehType
  Num1 = DCCustREc(1).FirstCar
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    Do Until cnt = 0
    'For cnt = Num1 To Num2
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).Active = "Y" Then
      GCode$ = Str$(cnt)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).DecalCat)
      GCode$ = GCode$ + "~" + Using$("$###.##", DCVRec(1).Fee)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).makemodel)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).StateTag)
      GCode$ = GCode$ + "~" + Num2Date$(DCVRec(1).ExpireDate)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Sticker)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).valid)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Desc)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).Notes)
      GCode$ = GCode$ + "~" + QPTrim$(DCVRec(1).PBFlag)
      dcnt = dcnt + 1
      Print #DCRpt, ToPrint$ + "~" + GCode$
      GCode$ = ""
    
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
  End If
    If dcnt <= 0 Then
        GCode$ = "0~  ~ ~No Vehicles to Display ~No Vehicles ~ ~ ~ ~ ~ ~ "
        ToPrint$ = ToPrint$ + "~" + GCode$
        Print #DCRpt, ToPrint$
        
    End If
  
  
  Close
  Load frmLoadingRpt
  'frmLoadingRpt.setwherefrom frmUBCustMenu
  ARptDCCustInfo.txtDate = Now
  ARptDCCustInfo.txtTown = TownName$
  ARptDCCustInfo.GetName ReportFile$
  ARptDCCustInfo.startrpt
  Else
  Print #DCRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
  Print #DCRpt, Tab(30); Title$
  Print #DCRpt, Now
  Print #DCRpt, TownName$
  Print #DCRpt, Dash80$
  Print #DCRpt,
  Print #DCRpt, "Cust #: "; QPTrim$(DCCustREc(1).CUSTNUMB);
  Print #DCRpt, Tab(50); "Search Name: "; QPTrim$(DCCustREc(1).SORTNAME)
  Print #DCRpt, "Customer Name: "; QPTrim$(DCCustREc(1).BILLNAME);
  Print #DCRpt, Tab(50); "Date Opened: "; Num2Date(DCCustREc(1).DATEOPED)
  Print #DCRpt, "Address: "; QPTrim$(DCCustREc(1).ADDRESS1)
  Print #DCRpt, Tab(10); QPTrim$(DCCustREc(1).ADDRESS2);
  Print #DCRpt, Tab(50); "Account Balance: "; Using$("$###,###,###.##", DCCustREc(1).AcctBal)
  Print #DCRpt, Tab(10); QPTrim$(DCCustREc(1).City); " "; QPTrim$(DCCustREc(1).State); " "; QPTrim$(DCCustREc(1).ZIPCODE)
  Print #DCRpt,
  Print #DCRpt, "DriverLic#: "; QPTrim$(DCCustREc(1).DRVLIC);
  Select Case DCCustREc(1).CASHONLY
  Case "N", " "
    Print #DCRpt, Tab(50); "  Cash Only: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "  Cash Only: "; "Yes"
  End Select
  Print #DCRpt, "SocSec#: "; QPTrim$(DCCustREc(1).SOSEC);
  Select Case DCCustREc(1).resident
  Case "N", " "
    Print #DCRpt, Tab(50); "Residential: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "Residential: "; "Yes"
  End Select
  Print #DCRpt, "Home Phone: "; QPTrim$(DCCustREc(1).HPHONE);
  Select Case DCCustREc(1).Owner
  Case "N", " "
    Print #DCRpt, Tab(50); "      Owner: "; "No"
  Case Else
    Print #DCRpt, Tab(50); "      Owner: "; "Yes"
  End Select
  Print #DCRpt, "Work Phone: "; QPTrim$(DCCustREc(1).WPHONE);
  If DCCustREc(1).FirstCar > 0 Then
    Print #DCRpt, Tab(45); "Vehicles on File: "; "Yes"
  Else
    Print #DCRpt, Tab(45); "Vehicles on File: "; "No"
  End If
  
  Print #DCRpt,
  Print #DCRpt, "-------------------------- Vehicle Information ----------------------"
  LineCnt = 20
  ReDim DCVRec(1) As DCVehType
  Num1 = DCCustREc(1).FirstCar
  If Num1 > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    cnt = Num1
    Do Until cnt = 0
    Get DCvFile, cnt, DCVRec(1)
    If DCVRec(1).Active = "Y" Then
      If LineCnt >= MaxLines Then
          Print #DCRpt, FF$
          Print #DCRpt, Tab(30); Title$
          Print #DCRpt, Now
          Print #DCRpt, TownName$
          Print #DCRpt, Dash80$
          Print #DCRpt,
          Print #DCRpt, "Cust #: "; QPTrim$(DCCustREc(1).CUSTNUMB);
          Print #DCRpt, Tab(50); "Search Name: "; QPTrim$(DCCustREc(1).SORTNAME)
          Print #DCRpt, "Customer Name: "; QPTrim$(DCCustREc(1).BILLNAME)
          Print #DCRpt, "Continued  ---------------- Vehicle Information ----------------------"
          LineCnt = 8
        End If
      Print #DCRpt, "Category   Fee        Sticker#       Expires      Valid    P/B"
      Print #DCRpt, "--------   ---        --------       -------      -----    ---"
      Print #DCRpt, QPTrim$(DCVRec(1).DecalCat); Tab(12); Using$("$###.##", DCVRec(1).Fee);
      Print #DCRpt, Tab(22); QPTrim$(DCVRec(1).Sticker); Tab(36); Num2Date$(DCVRec(1).ExpireDate);
      Print #DCRpt, Tab(52); QPTrim$(DCVRec(1).valid); Tab(60); QPTrim$(DCVRec(1).PBFlag)
      Print #DCRpt, "   Make/Model                            Vin#/Desc"
      Print #DCRpt, "   ----------                            ---------"
      Print #DCRpt, Tab(3); QPTrim$(DCVRec(1).makemodel); Tab(40); QPTrim$(DCVRec(1).Desc)
      Print #DCRpt, "   State Lic#                            Notes"
      Print #DCRpt, "   ----------                            -----"
      Print #DCRpt, Tab(3); QPTrim$(DCVRec(1).StateTag); Tab(40); QPTrim$(DCVRec(1).Notes)
      LineCnt = LineCnt + 9
      dcnt = dcnt + 1
    End If
      cnt = DCVRec(1).NextRec
    Loop 'Next
    Close DCvFile
  End If
    If dcnt <= 0 Then
        Print #DCRpt, "No Vehicles to Display**************"
        LineCnt = LineCnt + 1
    End If
 
  Print #DCRpt,
  Print #DCRpt, Dash80$
  Print #DCRpt, Chr$(12)

  Close

  ViewPrint ReportFile$, Title$
  KillFile ReportFile$
  End If
End Sub

Public Sub PrintTRListScreen()
  Unload frmTRDispList
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "UBTRlist.rpt", "Customer Transaction List"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "UBTRlist.rpt"
    ARptLineRpt.startrpt
  End If
End Sub
Public Sub PrintTRDetlScreen()
  Unload frmTRDetail
  Unload frmTRDispList
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "UBTRDetl.RPT", "Customer Detail Transaction"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "UBTRDetl.RPT"
    ARptLineRpt.startrpt
  End If
End Sub
Public Sub PrintConsmpScreen()
  Unload frmRptConsumpHist
  frmReportOpt.Show 1
  If rptopt = 2 Then
    ViewPrint "UBCnHist.RPT", "Customer Consumption History List"
  ElseIf rptopt = 1 Then
    Load frmLoadingRpt
    ARptLineRpt.GetName "UBCnHist.RPT"
    ARptLineRpt.startrpt
  End If
End Sub
Public Sub TXLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "TaxLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$; AcctLogFileName = "TaxLog.dat"
  Close AcctLogFile
End Sub
Public Static Sub DCLog(Text$)
  Dim Today As String, TheTime As String
  Dim AmPm As String, Hour As String
  Dim ThisHour As Integer, LogFile As Integer
  
  Today$ = Date$
  Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    ThisHour = Val(Mid$(TheTime$, 2, 1))
  Else
    ThisHour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case ThisHour
  Case Is > 11
    ThisHour = ThisHour - 12
    If ThisHour = 0 Then ThisHour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case ThisHour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(Str$(ThisHour))
    Case Else
      Hour$ = QPTrim$(Str$(ThisHour))
  End Select
  TheTime$ = Hour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open UBPath$ + "DCLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "CM/DC: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub
Public Function FillCatCMBO(x As fpCombo)
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumOFDCCatRecs As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  ghandle = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
  NumOFDCCatRecs = LOF(ghandle) \ DCCatCodeRecLen
  x.Row = 0
  For cnt = 1 To NumOFDCCatRecs
    Get #ghandle, cnt, DCCatCodeRec
    If DCCatCodeRec.InactiveFlag <> "Y" Then
      x.AddItem Str$(cnt) & Chr$(9) & QPTrim$(DCCatCodeRec.CATCODE) & Chr$(9) & DCCatCodeRec.CODEDESC
    Else
      x.AddItem Str$(cnt) & Chr$(9) & QPTrim$(DCCatCodeRec.CATCODE) & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
End Function
  
Public Sub OpenDCCustFile(NumOfDCRecs, DCFile)
  Dim DCCustRecLen As Integer
  Close DCFile
  ReDim DCCustREc(1) As DCCustRecType
  DCCustRecLen = Len(DCCustREc(1))
  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDcRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP
End Sub
Public Sub OpenDCCustIdxFile(NumOfDCIdxRecs, DCIdxFile)
  Dim DCCustIdxRecLen As Integer
  Close DCIdxFile
  ReDim DCCustIdxRec(1) As DCCustIDXRecType
  DCCustIdxRecLen = Len(DCCustIdxRec(1))
  DCIdxFile = FreeFile
  Open "DCCUST.IDX" For Random Access Read Write Shared As DCIdxFile Len = DCCustIdxRecLen
  NumOfDCIdxRecs = LOF(DCIdxFile) \ DCCustIdxRecLen
End Sub
Public Function DCCustCnt()
  Dim DCCustRecLen As Integer, DCFile As Integer, NumOfDCRecs As Long
  DCCustCnt = False

  ReDim tmpCustRec(1) As DCCustRecType
  DCCustRecLen = Len(tmpCustRec(1))

  DCFile = FreeFile
  Open "DCCUST.DAT" For Random Access Read Write Shared As DCFile Len = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  Close DCFile

  DCCustCnt = NumOfDCRecs

  Erase tmpCustRec

End Function
Public Sub LoadDCSetUpFile(dcSetUpRec() As DCSetupType, DCSetuplen)
  Dim Handle As Integer
  DCSetuplen = Len(dcSetUpRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "DCSetUP.dat") Then
    Handle = FreeFile
    Open UBPath$ + "DCSetUP.dat" For Random Shared As Handle Len = DCSetuplen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, dcSetUpRec(1)
    End If
    Close Handle
  End If
End Sub
Public Function GetDefaultDCLookUP%()
  ReDim dcSetUpRec(1) As DCSetupType
  Dim RecLen As Integer
  LoadDCSetUpFile dcSetUpRec(), RecLen
  GetDefaultDCLookUP = Val(dcSetUpRec(1).DefLook)
  Erase dcSetUpRec
End Function
Public Function StandBy(HoldIt As Long)
  Dim holdcnt As Long
  For holdcnt = 1 To HoldIt
  'this will just take a sec
  Next
End Function

Public Sub CheckHasTaxes(ByRef intHasTaxes As Integer)
    'Dim zz As String * 800
    Dim tmpHandle As Integer
    Dim txFileLen As Integer
    intHasTaxes = 0  'default no taxes
    On Error GoTo NoTaxExit
    
    tmpHandle = FreeFile
    Open TaxSetupName For Input As tmpHandle
    txFileLen = LOF(tmpHandle)
    Close #tmpHandle

    If txFileLen > 900 Then
       intHasTaxes = 2
    Else
       intHasTaxes = 1
    End If

NoTaxExit:
Close tmpHandle
On Error GoTo 0

End Sub

