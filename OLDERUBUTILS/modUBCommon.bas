Attribute VB_Name = "modUBCommon"
Option Explicit
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
Public Function uRound#(N#)
  uRound# = Int(N# * 100 + 0.5) / 100
End Function

'Public Function OKDeleteCust(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer, TotalBalance As Double
'  Dim M1 As String, M2 As String
'  Dim UBCustRecLen As Integer, UBCustF As Integer
'  If Recno& > 0 Then
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'  UBCustF = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'  Get UBCustF, Recno&, UBCustRec(1)
'  Close UBCustF
'
'  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'  If TotalBalance# <> 0 Then
'    UBLog "NODELETE:" + Str$(Recno&) + " BAL:" + Str$(TotalBalance#)
'    M1$ = "This account HAS A BALANCE"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    OKDeleteCust = False
'  ElseIf UBCustRec(1).DepositAmt <> 0 Then
'    UBLog "NODELETE:" + Str$(Recno&) + " DEP:" + Str$(UBCustRec(1).DepositAmt)
'    M1$ = "This account HAS A DEPOSIT"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    OKDeleteCust = False
'  ElseIf UBCustRec(1).Status <> "I" Then
'    UBLog "NODELETE:" + Str$(Recno&) + " NOT INACTIVE"
'    M1$ = "This account IS NOT INACTIVE"
'    M2$ = "CAN NOT DELETE THIS ACCOUNT!"
'    OKDeleteCust = False
'  Else
'    OKDeleteCust = True
'  End If
'  If OKDeleteCust = False Then
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    FntSize = frmMsgDialog.Label(1).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR!"
'    MsgText(1) = ""
'    MsgText(2) = M1$
'    MsgText(3) = ""
'    MsgText(4) = M2$
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'
'  End If
'End If
'End Function
Public Function ChkBillFile%()
  Dim OKFlag As Boolean, RecLen As Integer, FHand As Integer
  Dim NumOfRec As Long
  OKFlag = True 'assume all is well

  ReDim BillRec(1) As UBTransRecType
  RecLen = Len(BillRec(1))

  FHand = FreeFile
  Open UBBillsFile For Random Shared As FHand Len = RecLen
  NumOfRec& = LOF(FHand) \ RecLen
  Close FHand

  If NumOfRec& = 0 Then
    Kill UBBillsFile
    OKFlag = False
  End If

  ChkBillFile% = OKFlag

  Erase BillRec
End Function
'Public Function OKFinalCust(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer, TotalBalance As Double
'  Dim M1 As String, M2 As String
'  Dim UBCustRecLen As Integer, UBCustF As Integer
'  If Recno& > 0 Then
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'  UBCustF = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'  Get UBCustF, Recno&, UBCustRec(1)
'  Close UBCustF
'
'  If UBCustRec(1).Status <> "A" Then
'    UBLog "NOFinal:" + Str$(Recno&) + " NOT ACTIVE"
'    M1$ = "This account IS NOT ACTIVE"
'    M2$ = "CAN NOT SET THIS ACCOUNT TO FINAL!"
'    OKFinalCust = False
'  Else
'    OKFinalCust = True
'  End If
'  If OKFinalCust = False Then
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(3).FontSize
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    FntSize = frmMsgDialog.Label(1).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR!"
'    MsgText(1) = ""
'    MsgText(2) = M1$
'    MsgText(3) = ""
'    MsgText(4) = M2$
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'
'  End If
'End If
'End Function
'Public Function OKApplyDep(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  Dim UBCustRecLen As Integer, UBCustF As Integer
'  UBCustRecLen = Len(UBCustRec(1))
'  If Recno& > 0 Then
'    UBCustF = FreeFile
'    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'    Get UBCustF, Recno&, UBCustRec(1)
'    Close UBCustF
'
'    If UBCustRec(1).DepositAmt <= 0 Then
'      'OK = MsgBox%("UB", "NODPOSIT")
'      OKApplyDep = False
'    Else
'      OKApplyDep = True
'    End If
'    If OKApplyDep = False Then
'      frmMsgDialog.RetLabel = "-2"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      FntSize = frmMsgDialog.Label(1).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR!"
'      MsgText(1) = ""
'      MsgText(2) = "NO DEPOSIT"
'      MsgText(3) = ""
'      MsgText(4) = "This Account Has NO Deposit on File"
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'    End If
'
'  End If
'End Function
'Public Function OKDepRefund(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  Dim UBCustRecLen As Integer, UBCustF As Integer
'  UBCustRecLen = Len(UBCustRec(1))
'  If Recno& > 0 Then
'    UBCustF = FreeFile
'    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'    Get UBCustF, Recno&, UBCustRec(1)
'    Close UBCustF
'
'    If UBCustRec(1).DepositAmt <= 0 Then
'      'OK = MsgBox%("UB", "NODPOSIT")
'      OKDepRefund = False
'    Else
'      OKDepRefund = True
'    End If
'    If OKDepRefund = False Then
'      frmMsgDialog.RetLabel = "-2"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      FntSize = frmMsgDialog.Label(1).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR!"
'      MsgText(1) = ""
'      MsgText(2) = "NO DEPOSIT"
'      MsgText(3) = ""
'      MsgText(4) = "This Account Has NO Deposit on File"
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'    End If
'
'  End If
'End Function

'Public Function OKDepCreditAdj(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer, M1 As String, M2 As String, TotalBalance As Double, UBFile As Integer
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  Dim UBCustRecLen As Integer, UBCustF As Integer, TNum As Long, UBTranRecLen As Integer
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  If Recno& > 0 Then
'    UBCustF = FreeFile
'    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'    Get UBCustF, Recno&, UBCustRec(1)
'    TNum = UBCustRec(1).LastTrans
'    Close UBCustF
'    UBFile = FreeFile
'    Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
'    If TNum& > 0 Then
'      Get UBFile, TNum&, UBTranRec(1)
'    End If
'    Close UBFile
'    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
'    If UBCustRec(1).Status = "A" Then
'      M1$ = "Active Account"
'      M2$ = "This Account Is Active"
'      OKDepCreditAdj = False
'    ElseIf TNum& <= 0 Then
'      M1$ = "Missing Transactions"
'      M2$ = "NO Transactions for this Account"
'      OKDepCreditAdj = False
'    ElseIf UBTranRec(1).TransType <> TranAppliedDeposit Then
'      M1$ = "Missing Applied Deposit Transaction"
'      M2$ = "Last Trans NOT An Applied Deposit"
'      OKDepCreditAdj = False
'    ElseIf UBCustRec(1).DepositAmt <> 0 Then
'      M1$ = "Invalid Selection"
'      M2$ = "This Account Has A Deposit on File"
'      OKDepCreditAdj = False
'    ElseIf TotalBalance# >= 0 Then
'      M1$ = "Balance Not Credit"
'      M2$ = "No Credit Account Balance"
'      OKDepCreditAdj = False
'    Else
'      OKDepCreditAdj = True
'    End If
'    If OKDepCreditAdj = False Then
'      frmMsgDialog.RetLabel = "-2"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      FntSize = frmMsgDialog.Label(1).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR!"
'      MsgText(1) = ""
'      MsgText(2) = M1$
'      MsgText(3) = ""
'      MsgText(4) = M2$
'      MsgText(5) = ""
'      GetOKorNot MsgText(), True
'    End If
'
'  End If
'End Function
'Public Function OKDepReverse(Recno&)
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer, M1 As String, M2 As String, TotalBalance As Double, UBFile As Integer
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1 To 2) As NewUBCustRecType
'  Dim dcnt As Long, FoundCM As Long, FoundCnt As Long, M3 As String
'  Dim UBCustRecLen As Integer, UBCustF As Integer, PrevTranRec As Long, UBTranRecLen As Integer
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  If Recno& > 0 Then
'    UBCustF = FreeFile
'    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
'    Get UBCustF, Recno&, UBCustRec(1)
'    PrevTranRec& = UBCustRec(1).LastTrans
'    UBFile = FreeFile
'    Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
'
'    If PrevTranRec& <= 0 Then
'      M1$ = "Missing Transactions"
'      M2$ = "NO Transactions for this Account"
'      OKDepReverse = False
'    ElseIf UBCustRec(1).DepositAmt = 0 Then
'      M1$ = "Invalid Selection"
'      M2$ = "This Account Has No Deposit on File"
'      OKDepReverse = False
'    Else
'      If PrevTranRec& > 0 Then
'        Do While PrevTranRec& > 0
'          dcnt = dcnt + 1
'          Get UBFile, PrevTranRec&, UBTranRec(1)
'           If UBTranRec(1).TransType = TranDepositPayment Or UBTranRec(1).TransType = TranDepositPayment + 100 Then
'            If UBTranRec(1).VoidFlag = True Then
'             'just skip to next
'            Else
'             If UBTranRec(1).FromCMFlag = True Then
'              FoundCM = FoundCM + 1
'             Else
'              FoundCnt = FoundCnt + 1
'             End If
'            End If
'           End If
'          PrevTranRec& = UBTranRec(1).PrevTrans
'        Loop
'      End If
'      If FoundCnt > 0 And FoundCM = 0 Then
'       OKDepReverse = True
'      ElseIf FoundCM > 0 Then
'       M1$ = "Deposit Payment Taken in"
'       M2$ = "Cash Management."
'       M3$ = "Must be voided thru CM."
'       OKDepReverse = False
'     Else
'      M1$ = "Deposit Payment Voided"
'      M2$ = "Can Not Be Voided Again."
'      OKDepReverse = False
'     End If
'
'  End If
'  Close UBFile
'  Erase UBTranRec, UBCustRec
'  DoEvents
'
'    If OKDepReverse = False Then
'      frmMsgDialog.RetLabel = "-2"
'      FntSize = frmMsgDialog.Label(3).FontSize
'      frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'      FntSize = frmMsgDialog.Label(1).FontSize
'      frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'      MsgText(0) = "ERROR!"
'      MsgText(1) = ""
'      MsgText(2) = M1$
'      MsgText(3) = ""
'      MsgText(4) = M2$
'      MsgText(5) = M3$
'      GetOKorNot MsgText(), True
'    End If
'
'  End If
'End Function
Public Sub UPDateOK()
  frmDataUpdated.Show vbModal
End Sub

'!!! Procedures below Needed for reports!!! Mark with!!!
'Make sure to check w/Dale  PS
'!!! Added Round on 4-17-03
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function
'loads Work Order Defaults into fpcombos
'Public Sub GetWOList(X As fpCombo)
'  Dim cnt As Long, NumWOs As Long
'  Dim WorkOrderDefLen As Integer
'  Dim UBWrkOrdD As Integer
'
'  Dim WorkOrderDef As WorkOrderDefType
'  WorkOrderDefLen = Len(WorkOrderDef)
'
'  UBWrkOrdD = FreeFile
'  Open UBPath$ + "UBWODef.DAT" For Random Shared As UBWrkOrdD Len = WorkOrderDefLen
'  NumWOs = LOF(UBWrkOrdD) \ WorkOrderDefLen
'  For cnt = 1 To NumWOs
'    Get UBWrkOrdD, cnt, WorkOrderDef
'      If WorkOrderDef.Deleted <> True Then
'        X.InsertRow = Str(cnt) & Chr$(9) & QPTrim(WorkOrderDef.WOType)
'      End If
'  Next
'  Close
'End Sub

'!!! populates the combo box with revenues
Public Function FillRevList(x As fpCombo)
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  Dim cnt As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  'X.AddItem "All Revenues"
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
Public Function GetZipEDigit$(Zip$)
  Dim ZipLen As Integer, ZipVal As Integer, DashPos As Integer
  Dim cnt As Integer, Dif As Double
  ZipLen = Len(Zip$)
  ZipVal = 0

  DashPos = InStr(Zip$, "-")
  Do While DashPos
    Zip$ = Left$(Zip$, DashPos - 1) + Mid$(Zip$, DashPos + 1)
    DashPos = InStr(Zip$, "-")
  Loop

  For cnt = 1 To ZipLen
    ZipVal = ZipVal + Val(Mid$(Zip$, cnt, 1))
  Next

  If ZipVal Mod 10 > 0 Then
    Dif = 10 - (ZipVal Mod 10)
  Else
    Dif = 0
  End If
  GetZipEDigit$ = QPTrim$(Str$(Dif))

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
Public Sub LoadUBBillSetUpFile(UBBillSetUpRec() As UBBillSetupType, UBBillSetuplen)
  Dim Handle As Integer
  UBBillSetuplen = Len(UBBillSetUpRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "UBBilSET.DAT") Then
    Handle = FreeFile
    Open UBPath$ + "UBBilSET.DAT" For Random Shared As Handle Len = UBBillSetuplen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, UBBillSetUpRec(1)
    End If
    Close Handle
  End If
End Sub

Public Function Exist(FileName$)
'  On Local Error Resume Next
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
'  If Err Then
'    FileName$ = ""
'  End If
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
'  On Local Error GoTo 0
End Function

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
  Dim CHandle As Integer, cnt As Long, IHandle As Integer
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
  
'  Load frmInfo
'  frmInfo.Label1 = "Sorting. . ."
'  DoEvents
'  frmInfo.Show
  DoEvents
  SeqQSort SequenceIndex(), 1, NumCustRecs&
'  Unload frmInfo
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
'Public Sub MakeMowZipCodeIndex(IndexText$)
'  FrmShowPctComp.Label1 = "Reading Customer Information."
'  FrmShowPctComp.CmdCancel.Enabled = False
'  FrmShowPctComp.Show '1, Parent
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  Dim CustRecLen As Integer, IndexRecLen As Integer
'  Dim CHandle As Integer, cnt As Integer, IHandle As Integer
'  Dim NumCustRecs As Long, Prec As Long, NumOfBillRec As Long
'  Dim BCnt As Long
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
'
'  CHandle = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'
'  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
'  For BCnt = 1 To NumOfBillRec
'    Get CHandle, BCnt, UBCustRec(1)
'    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    ZipIndex(BCnt).RecNum = BCnt
'    FrmShowPctComp.ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
'  Next
'  Close
''  Load frmInfo
''  frmInfo.Label1 = "Sorting. . ."
''  DoEvents
''  frmInfo.Show
'  DoEvents
'  ZipQSort ZipIndex(), 1, NumOfBillRec
''  Unload frmInfo
'  DoEvents
'
'  FrmShowPctComp.Label1 = "Writing Index Records."
'  FrmShowPctComp.CmdCancel.Enabled = False
'  FrmShowPctComp.Show
'
' KillFile TempIndexName
'  IHandle = FreeFile
'  Open TempIndexName For Output As IHandle
'  Close IHandle
'
'  IHandle = FreeFile
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For cnt = 1 To NumOfBillRec
'    Prec& = ZipIndex(cnt).RecNum
'    Put IHandle, cnt, Prec&
'    FrmShowPctComp.ShowPctComp cnt, NumOfBillRec               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ZipIndex
'
'End Sub
'Public Sub MakeZipCodeIndex(IndexText$)
''Removed all rest of code
'  Call MakeMowZipCodeIndex(IndexText$)
'
'End Sub
'
''For Mail Lables
'Public Sub MakePostalIndex(IndexText$)
'  FrmShowPctComp.Label1 = "Reading Customer Information."
'  FrmShowPctComp.CmdCancel.Enabled = False
'  FrmShowPctComp.Show '1, Parent
'
'  Dim CustRecLen As Integer, IndexRecLen As Integer
'  Dim CHandle As Integer, cnt As Integer, IHandle As Integer
'  Dim NumCustRecs As Long, Prec As Long
'  Dim BCnt As Long
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumCustRecs = GetNumOfCust
'
'  ReDim PostalIndex(1 To NumCustRecs) As UBPostalIndexType
'  IndexRecLen = Len(PostalIndex(1))
'
'  CHandle = FreeFile
'  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'  For cnt = 1 To NumCustRecs
'    Get CHandle, cnt, UBCustRec(1)
'    PostalIndex(cnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    RSet PostalIndex(cnt).Route = QPTrim$(UBCustRec(1).POSTRTE)
'    PostalIndex(cnt).RecNum = cnt
'    FrmShowPctComp.ShowPctComp cnt, NumCustRecs                'show user percentage complete
'  Next
'
'  Close CHandle
'  Load frmInfo
'  frmInfo.Label1 = "Sorting. . ."
'  DoEvents
'  frmInfo.Show
'  DoEvents
'  PostalQSort PostalIndex(), 1, NumCustRecs
'  Unload frmInfo
'  DoEvents
'
'  FrmShowPctComp.Label1 = "Writing Index Records."
'  FrmShowPctComp.CmdCancel.Enabled = False
'  FrmShowPctComp.Show
'
'  IHandle = FreeFile
'
'  'FCreate TempIndexName
'  KillFile TempIndexName
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For cnt = 1 To NumCustRecs
'    Prec& = PostalIndex(cnt).RecNum
'    Put IHandle, cnt, Prec&
'    FrmShowPctComp.ShowPctComp cnt, NumCustRecs                'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, PostalIndex
'End Sub
'Function returns True if a customer has been deleted.
Public Function IsDeleted%(AcctNum&)
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim Handle As Integer
  Dim UBCustRecLen As Integer
  
  UBCustRecLen = Len(UBCustRec(1))
  Handle = FreeFile
  Open UBCustFile For Random Shared As Handle Len = UBCustRecLen
  Get Handle, AcctNum&, UBCustRec(1)
  Close Handle
  
  If UBCustRec(1).DelFlag <> 0 Then
    IsDeleted% = True
  Else
    IsDeleted% = False
  End If
  Erase UBCustRec

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

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String, Optional HideF7btn As Boolean)
 ' frmLoadingRpt.Show
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
  If HideF7btn Then
    frmViewPrint.cmdPrnScn.Enabled = False
  End If
  DoEvents
' Unload frmLoadingRpt
  DoEvents
  frmViewPrint.Show vbModal
End Sub
'Public Sub ViewPrintM(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String, Optional HideF7btn As Boolean)
' ' frmLoadingRpt.Show 1
' 'not using loadingrpt form only diff between this and regular viewprint
' 'the problem was all modal forms
'  DoEvents
'  frmViewPrint.ReportName = ReportFile$
'  frmViewPrint.Caption = Title
'  frmViewPrint.PgNum = PgNum
'  If ForceSBar Then
'    frmViewPrint.fpMemo1.ScrollBars = BothFixed
'  Else
'    frmViewPrint.fpMemo1.ScrollBars = BothAuto
'  End If
'  If Algn Then
'    frmViewPrint.cmdAlignment.Enabled = True
'    frmViewPrint.AlignRpt = AlgnRptfile$
'  Else
'    frmViewPrint.cmdAlignment.Enabled = False
'  End If
'  If HideF7btn Then
'    frmViewPrint.cmdPrnScn.Enabled = False
'  End If
' ' DoEvents
' 'Unload frmLoadingRpt
'  DoEvents
'  frmViewPrint.Show vbModal
'End Sub

Public Function GetDefaultLookUP%()
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  GetDefaultLookUP = Val(UBSetUpRec(1).DefLook)
  Erase UBSetUpRec
End Function

'Function to format the Book part of a location number
Public Function FmtBook$(Book$)
  Dim BookLen As Integer
  
  Book$ = QPTrim$(Book$)
  BookLen = Len(Book$)
  
  Select Case BookLen
  Case 0
'Per Dale to fix autofill problem change from 00 to spaces
    'FmtBook$ = "00"
    FmtBook$ = "  "
  Case 1
    FmtBook$ = "0" + Book$
  Case Else
    FmtBook$ = Book$
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
  'Per Dale to fix autofill problem change from 0's to 6 spaces
   ' FmtSeqN$ = "000000"
    FmtSeqN$ = "      "
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

'Public Sub UBTerminate()
'  Dim UBFrmCnt As Integer
'  ' Loop through the forms collection and unload each form.
'  On Local Error Resume Next
'  UBLog "UB Exited: "
'  ClearInUse PWcnt
'  If DebugMode = False Then
'    Shell "CitiPak.exe", vbMaximizedFocus
'  End If
'  DoTheTime
'  DoEvents
'  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
'    DoEvents
'    Unload Forms(UBFrmCnt)
'  Next
'  End
'End Sub
'Public Sub CitiTerminate()
'  Dim UBFrmCnt As Integer
'  ' Loop through the forms collection and unload each form.
'  ClearInUse PWcnt
'  For UBFrmCnt = Forms.Count - 1 To 0 Step -1
'    Unload Forms(UBFrmCnt)
'  Next
'  DoEvents
'  End
'End Sub

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
  Open UBPath$ + "UButilLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "UButil: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub
'Public Sub ResetProRates()
'
'  Load frmNoOperatorsWarning
'  frmNoOperatorsWarning.Label(5) = "CONTINUE WITH RESET PRORATES?"
'  frmNoOperatorsWarning.Show vbModal
'
'  If Not DoItFlag Then
'    UBLog "ABORTED: Reset Prorate Percentages"
'    GoTo ExitResetProRates
'  End If
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  Dim CustRecLen As Integer
'  Dim UBFile As Integer, CCnt As Long
'  Dim NumOfCRecs As Long
'
'  FrmShowPctComp.Label1 = "Reading Customer Information."
'  FrmShowPctComp.AutoClose = "no"
'  FrmShowPctComp.CmdCancel.Enabled = False
'  FrmShowPctComp.Show '1, Parent
'  DoEvents
'
'  CustRecLen = Len(UBCustRec(1))
'  'NumOfCRecs& = GetNumOfCust&
'  UBLog "BEGIN: Reseting Percentages"
'
'  UBFile = FreeFile
'  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = CustRecLen
'  NumOfCRecs = LOF(UBFile) \ CustRecLen
'
'  For CCnt = 1 To NumOfCRecs
'    Get UBFile, CCnt, UBCustRec(1)
'    UBCustRec(1).ProRatePCT = 100
'    Put UBFile, CCnt, UBCustRec(1)
'    FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs
'    'ShowPctComp
'  Next
'  Close
'  Erase UBCustRec
'  Unload FrmShowPctComp
'  UPDateOK
'
'ExitResetProRates:
'
'End Sub

Public Sub ReIndexSystem(PromptFlag%)

  UBLog " IN: Reindex Utility Files"
  Dim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBFile As Integer, BookHand As Integer
  Dim cnt As Long
  Dim NumOfRecs As Long
  Dim TmpBookSeq As String
  DoItFlag = False
  
  If PromptFlag% Then
    Load frmNoOperatorsWarning
    frmNoOperatorsWarning.Show vbModal
    If Not DoItFlag Then
      GoTo ExitReindex
    End If
  End If

  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.AutoClose = "no"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

  UBCustRecLen = Len(UBCustRec(1))              'Length of Cust Record Structure
  UBLog "BEGIN: Customer Name Reindex"
  UBFile = FreeFile
  Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  ReDim IdxBuff(1 To NumOfRecs&) As nUBCustReIndexRecType

  For cnt = 1 To NumOfRecs&
    Get UBFile, cnt, UBCustRec(1)
    IdxBuff(cnt).SearchName = UBCustRec(1).SEARCH
    If UBCustRec(1).DelFlag Then
      IdxBuff(cnt).DelFlag = "Y"
    Else
      IdxBuff(cnt).DelFlag = ""
    End If
    IdxBuff(cnt).Status = UBCustRec(1).Status
    IdxBuff(cnt).RecNum = cnt
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
  Next

  Close UBFile

  FrmShowPctComp.Label1 = "Sorting Customer Names"
  NameQSort IdxBuff(), 1, NumOfRecs&
  FrmShowPctComp.Label1 = "Writing Customer Index"
  KillFile "UBCUSTNM.IDX"
  UBFile = FreeFile
  Open UBPath + "UBCUSTNM.IDX" For Random Shared As UBFile Len = 4
  For cnt = 1 To NumOfRecs&
    Put UBFile, cnt, IdxBuff(cnt).RecNum
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
  Next
  Close UBFile
  FrmShowPctComp.Label1 = "Writing Customer Search Data"
  KillFile "UBCUSTSN.DAT"
  UBFile = FreeFile
  Open UBPath + "UBCUSTSN.DAT" For Random Shared As UBFile Len = Len(IdxBuff(1))
  For cnt = 1 To NumOfRecs&
    Put UBFile, cnt, IdxBuff(cnt)
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
  Next
  Close UBFile

  Erase IdxBuff
  UBLog "FINISH: Customer Name Reindex"

  FrmShowPctComp.Label1 = "Reading Location Information"
  UBLog "BEGIN: Book\Sequence Reindex"

  UBFile = FreeFile
  Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  ReDim LIdxBuff(1 To NumOfRecs&) As UBLocaReIndexRecTypeVB

  For cnt = 1 To NumOfRecs&
    Get UBFile, cnt, UBCustRec(1)
    TmpBookSeq = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
    LIdxBuff(cnt).BookSEQNUMB = TmpBookSeq
    LIdxBuff(cnt).RecNum = cnt
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
  Next

  Close UBFile

  FrmShowPctComp.Label1 = "Sorting Locations Names"
  LocQSort LIdxBuff(), 1, NumOfRecs&

  FrmShowPctComp.Label1 = "Writing Location Index"
  KillFile "UBCUSTBK.IDX"
  UBFile = FreeFile
  Open UBPath + "UBCUSTBK.IDX" For Random Shared As UBFile Len = 4
  For cnt = 1 To NumOfRecs&
    Put UBFile, cnt, LIdxBuff(cnt).RecNum
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
  Next
  Close UBFile

  UBLog "FINISH: Book\Sequence Reindex"
  ReDim BookSeq(1) As BookSeqRecType

  KillFile "UBOOKSEQ.DAT"
  UBLog "BEGIN: Rebuild Book\Sequence List"
  BookHand = FreeFile
  Open UBPath + "UBOOKSEQ.DAT" For Random Shared As BookHand Len = 4
  For cnt = 1 To NumOfRecs&
    BookSeq(1).BookSeq = Val(LIdxBuff(cnt).BookSEQNUMB)
    Put BookHand, cnt, BookSeq(1)
  Next
  Close BookHand
  UBLog "FINISH: Rebuild Book\Sequence List"

  Erase LIdxBuff, BookSeq, IdxBuff
  Erase UBCustRec
  Unload FrmShowPctComp
  If PromptFlag% Then
    UPDateOK
  End If
'  MsgBox "Done"

ExitReindex:
  UBLog "OUT: Reindex Utility Files" + CrLf$
  Exit Sub

End Sub
Public Sub Fixlastreadsndate() 'pulls last reading from last bill trans....
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer, Amount#
  Dim UBFile As Integer, UBTran As Integer, ReportFile$, UBRpt As Integer
  Dim NumOfCRecs As Long, NumOfTRecs As Long, TotalTrans#, Trans&
  Dim TRRecs As Long, RevAmts As Integer, MtrCnt As Integer
  Dim CCnt As Long, TotalCustBalance#
  Dim UBTransRec As UBTransRecType
  Dim UBCustRec As NewUBCustRecType
 
  UBCustRecLen = Len(UBCustRec)              'Length of Cust Record Structure
  UBTranRecLen = Len(UBTransRec)             'Length of Tran Record Structure

  UBTran = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  NumOfTRecs = LOF(UBTran) \ UBTranRecLen

  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfCRecs = LOF(UBFile) \ UBCustRecLen
'  ReportFile$ = UBPath$ + "UBunfix.RPT"
'  UBRpt = FreeFile
'  Open ReportFile$ For Output As UBRpt

  FrmShowPctComp.Label1 = "Redo last reads"
  FrmShowPctComp.Show
  For CCnt = 1 To NumOfCRecs
    FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs&
    For MtrCnt = 1 To 7
      UBCustRec.LocMeters(MtrCnt).PrevRead = 0
      UBCustRec.LocMeters(MtrCnt).CurRead = 0
      UBCustRec.LocMeters(MtrCnt).PastDate = 0
      UBCustRec.LocMeters(MtrCnt).CurDate = 0
      UBCustRec.LocMeters(MtrCnt).ReadFlag = ""
    Next
    Get UBFile, CCnt, UBCustRec
    Trans& = UBCustRec.LastTrans
      Do While Trans& > 0
      Get UBTran, Trans&, UBTransRec
        If UBTransRec.TransType = TranUtilityBill Then
          For MtrCnt = 1 To 7
            UBCustRec.LocMeters(MtrCnt).PrevRead = UBTransRec.PrevRead(MtrCnt)
            UBCustRec.LocMeters(MtrCnt).CurRead = UBTransRec.CurRead(MtrCnt)
            UBCustRec.LocMeters(MtrCnt).PastDate = UBTransRec.PrevDate
            UBCustRec.LocMeters(MtrCnt).CurDate = UBTransRec.ReadDate
            UBCustRec.LocMeters(MtrCnt).ReadFlag = ""
          Next
          Trans& = 0
          Put UBFile, CCnt, UBCustRec
          Exit Do
        Else
          Trans& = UBTransRec.PrevTrans
        End If
      Loop
    Next
End Sub
Public Sub SetAvgusetoONE()
  Dim UBCustRecLen As Integer
  Dim UBFile As Integer
  Dim NumOfCRecs As Long
  Dim MtrCnt As Integer
  Dim CCnt As Long
  Dim UBCustRec As NewUBCustRecType
 
  UBCustRecLen = Len(UBCustRec)              'Length of Cust Record Structure


  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfCRecs = LOF(UBFile) \ UBCustRecLen
'  ReportFile$ = UBPath$ + "UBunfix.RPT"
'  UBRpt = FreeFile
'  Open ReportFile$ For Output As UBRpt
  For CCnt = 1 To NumOfCRecs
  Get UBFile, CCnt, UBCustRec
    For MtrCnt = 1 To 7
      UBCustRec.LocMeters(MtrCnt).AvgUse = 1
      UBCustRec.LocMeters(MtrCnt).UseCnt = 1
    Next
   Put UBFile, CCnt, UBCustRec
  Next
  MsgBox "Done", vbOKOnly
End Sub

Public Sub RecalcUBCustBalances()
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer, Amount#
  Dim UBFile As Integer, UBTran As Integer, ReportFile$, UBRpt As Integer
  Dim NumOfCRecs As Long, NumOfTRecs As Long, TotalTrans#, Trans&
  Dim TRRecs As Long, RevAmts As Integer
  Dim CCnt As Long, TotalCustBalance#
  Dim UBTransRec As UBTransRecType
  Dim UBCustRec As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec)              'Length of Cust Record Structure
  UBTranRecLen = Len(UBTransRec)             'Length of Tran Record Structure

  UBTran = FreeFile
  Open UBPath + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  NumOfTRecs = LOF(UBTran) \ UBTranRecLen

  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfCRecs = LOF(UBFile) \ UBCustRecLen


  For CCnt = 1 To NumOfCRecs
    Get UBFile, CCnt, UBCustRec
      UBCustRec.PrevBalance = 0
      UBCustRec.CurrBalance = 0
      UBCustRec.DepositAmt = 0
      For RevAmts = 1 To 15
        UBCustRec.CurrRevAmts(RevAmts) = 0
      Next
    Put UBFile, CCnt, UBCustRec
  Next
  MsgBox "Balances reset to zero", vbOKOnly
'  FrmShowPctComp.Label1 = "Recalculating Balances"
'  FrmShowPctComp.Show
'  For CCnt = 1 To NumOfCRecs
'    FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs&
'    Get UBFile, CCnt, UBCustRec
'    For TRRecs& = 1 To NumOfTRecs
'      Get UBTran, TRRecs&, UBTransRec
'      If UBTransRec.CustAcctNo = CCnt Then
'        GoSub DotheTrans
'        UBTransRec.RunBalance = Round#(UBCustRec.PrevBalance + UBCustRec.CurrBalance)
'        Put UBTran, TRRecs&, UBTransRec
'      End If
''      For RevAmts = 1 To MaxRevsCnt
''        UBTransRec.RevAmt(RevAmts) = 0
''        UBTransRec.TaxAmt(RevAmts) = 0
''      Next
''      UBTransRec.Transamt = 0
'    Next
'    Put UBFile, CCnt, UBCustRec
'      UBCustRec.PrevBalance = 0
'      UBCustRec.CurrBalance = 0
'      For RevAmts = 1 To MaxRevsCnt
'        UBCustRec.CurrRevAmts(RevAmts) = 0
'        UBTransRec.RevAmt(RevAmts) = 0
'        UBTransRec.TaxAmt(RevAmts) = 0
'      Next
'      UBCustRec.DepositAmt = 0
'      UBTransRec.Transamt = 0
'
'  Next
'  Close
  FrmShowPctComp.Label1 = "Recalculating Balances"
  FrmShowPctComp.Show
    
    For TRRecs& = 1 To NumOfTRecs
    FrmShowPctComp.ShowPctComp TRRecs&, NumOfTRecs
      Get UBTran, TRRecs&, UBTransRec
      If UBTransRec.CustAcctNo > 0 And UBTransRec.CustAcctNo <= NumOfCRecs Then
      Get UBFile, UBTransRec.CustAcctNo, UBCustRec
      ''''If UBTransRec.CustAcctNo = 8046 Then Stop
        If UBCustRec.DelFlag = 0 Then
        GoSub DotheTrans
        UBTransRec.RunBalance = Round#(UBCustRec.PrevBalance + UBCustRec.CurrBalance)
        Put UBTran, TRRecs&, UBTransRec
        Put UBFile, UBTransRec.CustAcctNo, UBCustRec
        UBCustRec.PrevBalance = 0
        UBCustRec.CurrBalance = 0
       For RevAmts = 1 To 15
        UBCustRec.CurrRevAmts(RevAmts) = 0
        UBTransRec.RevAmt(RevAmts) = 0
        UBTransRec.TaxAmt(RevAmts) = 0
       Next
       UBCustRec.DepositAmt = 0
       UBTransRec.Transamt = 0
       End If
      End If
  Next
  Close
  MsgBox "Done", vbOKOnly
  Exit Sub
  
DotheTrans:
'
  Select Case UBTransRec.TransType
  Case TranUtilityBill, TranUtilityBill + 100
      UBCustRec.PrevBalance = Round#(UBCustRec.PrevBalance + UBCustRec.CurrBalance)
      UBCustRec.CurrBalance = UBTransRec.Transamt
      For RevAmts = 1 To 14
        UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) + UBTransRec.RevAmt(RevAmts) + UBTransRec.TaxAmt(RevAmts))
      Next
  Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
  'what
  Case TranBillPayment, TranBillPayment + 100
      If UBCustRec.PrevBalance <> 0 Then
        If UBTransRec.Transamt >= UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = 0
        ElseIf UBTransRec.Transamt < UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = Round#(UBCustRec.PrevBalance - UBTransRec.Transamt)
        End If
      End If
      For RevAmts = 1 To 14
        UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) - UBTransRec.RevAmt(RevAmts))
      Next
      TotalCustBalance# = 0
      For RevAmts = 1 To 14
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec.CurrRevAmts(RevAmts))
      Next
      UBCustRec.CurrBalance = Round#(TotalCustBalance# - UBCustRec.PrevBalance)
  Case TranPenaltyCharge
      UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance + UBTransRec.Transamt)
      For RevAmts = 1 To 14
        UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) + UBTransRec.RevAmt(RevAmts))
      Next
  Case TranAppliedDeposit
      For RevAmts = 1 To 14
        UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) - UBTransRec.RevAmt(RevAmts))
      Next
      TotalCustBalance# = 0
      For RevAmts = 1 To 14
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec.CurrRevAmts(RevAmts))
      Next
      UBCustRec.CurrBalance = Round#(TotalCustBalance#)
      If UBCustRec.PrevBalance > 0 Then
        If UBCustRec.DepositAmt >= UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = 0
        ElseIf UBCustRec.DepositAmt < UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = Round#(UBCustRec.PrevBalance - UBCustRec.DepositAmt)
          UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.PrevBalance)
        End If
      ElseIf UBCustRec.PrevBalance < 0 Then
        UBCustRec.PrevBalance = 0
      End If
      UBCustRec.DepositAmt = 0
'      Select Case UBCustRec.PrevBalance
'        Case 0
'          'UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.DepositAmt)
'        Case Is > 0
'          If UBCustRec.PrevBalance < UBCustRec.DepositAmt Then
'           UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.PrevBalance)
'           'UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.DepositAmt)
'          ' UBCustRec.PrevBalance = 0
'          Else
'
'           UBCustRec.PrevBalance = Round#(UBCustRec.PrevBalance - UBCustRec.DepositAmt)
'
'          End If
'        Case Is < 0
'          'UBCustRec.PrevBalance = 0
'         UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.PrevBalance)
'         'UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBCustRec.DepositAmt)
'        End Select
  Case TranDepositPayment, TranDepositPayment + 100
      UBCustRec.DepositAmt = UBCustRec.DepositAmt + UBTransRec.Transamt
  Case TranDraftPayment
      If UBCustRec.PrevBalance <> 0 Then
        If UBTransRec.Transamt >= UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = 0
        ElseIf UBTransRec.Transamt < UBCustRec.PrevBalance Then
          UBCustRec.PrevBalance = Round#(UBCustRec.PrevBalance - UBTransRec.Transamt)
        End If
      End If
      For RevAmts = 1 To 14
        UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) - UBTransRec.RevAmt(RevAmts))
      Next
      TotalCustBalance# = 0
      For RevAmts = 1 To 14
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec.CurrRevAmts(RevAmts))
      Next
      UBCustRec.CurrBalance = Round#(TotalCustBalance# - UBCustRec.PrevBalance)
  Case TranBeginBalance, TranBeginBalance + 100
    'top
  Case 9
    UBCustRec.DepositAmt = 0
  Case TranUpwardAdjustment
    For RevAmts = 1 To 14
      UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) + UBTransRec.RevAmt(RevAmts))
      UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance + UBTransRec.RevAmt(RevAmts))
    Next
  Case TranDownwardAdjustment
    For RevAmts = 1 To 14
      UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) - UBTransRec.RevAmt(RevAmts))
      UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance - UBTransRec.RevAmt(RevAmts))
    Next
  Case TranOverPayAdjustment
    For RevAmts = 1 To 14
      UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) + UBTransRec.RevAmt(RevAmts))
      UBCustRec.CurrBalance = Round#(UBCustRec.CurrBalance + UBTransRec.RevAmt(RevAmts))
    Next
  Case TranDepCreditRemoval
    For RevAmts = 1 To 14
      UBCustRec.CurrRevAmts(RevAmts) = Round#(UBCustRec.CurrRevAmts(RevAmts) + UBTransRec.RevAmt(RevAmts))
    Next
      UBCustRec.CurrBalance = 0
      UBCustRec.PrevBalance = 0
  Case TranDepPaymentVoid
    If UBCustRec.DepositAmt >= UBTransRec.Transamt Then
      UBCustRec.DepositAmt = Round#(UBCustRec.DepositAmt - UBTransRec.Transamt)
    Else
      UBCustRec.DepositAmt = 0
    End If
      Case Else
    'TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select

Return

ExitThisthing:
  'UBLog "OUT: Relink Transaction History" + CrLf$
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
      Mid$(Build, 48) = Using("#####.##", UBTranRec(1).Transamt, True)
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

Public Function CustHasMsg(Recno&)
  
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim MsgRec(1) As UBMessRecType
  Dim MsgLen As Integer, UBCustRecLen As Integer
  Dim UBFile As Integer, zz As Integer
  Dim NumMsgRec As Long, MRec As Long
  
  CustHasMsg = False
  
  MsgLen = Len(MsgRec(1))
  NumMsgRec& = FileSize&("UBMESAGE.DAT") / MsgLen

  UBCustRecLen = Len(UBCustRec(1))

  If Recno& > 0 Then
    UBFile = FreeFile
    Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
    Get UBFile, Recno&, UBCustRec(1)
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

Public Function GetOKorNot%(MsgText() As String, Optional OKOnly As Boolean, Optional ByVal NoFlash As Boolean, Optional ByVal Add2Font As Integer)
  Dim zz As Integer, RetValue As Integer
  If OKOnly Then
    frmMsgDialog.RetLabel = "-2"
  End If
  frmMsgDialog.Caption = MsgText(0)
  For zz = 1 To 5
    frmMsgDialog.Label(zz - 1) = MsgText(zz)
    If Add2Font > 0 Then
      frmMsgDialog.Label(zz - 1).FontSize = frmMsgDialog.Label(zz - 1).FontSize + Add2Font
    End If
  Next
  If NoFlash Then
    frmMsgDialog.Timer1.Enabled = False
  End If
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

Public Sub RateCodeErrScrn(RATECODE$)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "RATE CODE:  " + QPTrim$(RATECODE$)
    MsgText(3) = "Has an INVALID entry!"
    MsgText(4) = ""
    MsgText(5) = "Correct and Print Again."
    GetOKorNot MsgText(), True
End Sub

Public Static Function GetRevCharge#(RateTbl As UBRateTblRecType, TMeterConsp&, MeterMulti&)
  Dim MinBillAmt As Double, TAmt As Double, LastTblCnt As Integer
  Dim BCnt As Integer, MeterConsump As Long, UNITS As Long
  'STOP

  MinBillAmt# = RateTbl.MINAMT

  If MinBillAmt# < -1000000 Then
    MinBillAmt# = 0
    TAmt# = -1
    GoTo GotTAmt
  End If

'SunnyBeech 091701
  If TMeterConsp& <= RateTbl.MINUNITS Then
    TAmt# = 0
    GoTo GotTAmt
  End If

  LastTblCnt = 10
  For BCnt = 1 To 10
    If RateTbl.TblBreaks(BCnt).UNITAMT <= 0 Then
      LastTblCnt = BCnt - 1
      Exit For
    End If
  Next

  MeterConsump& = TMeterConsp&

  TAmt# = 0

  If LastTblCnt >= 2 Then
    If MeterConsump& >= RateTbl.TblBreaks(1).UNITS And MeterConsump& <= RateTbl.TblBreaks(2).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
      'special patch for cave junction
      If UNITS& = 0 Then
        UNITS& = 1
      End If
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    End If
  Else          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GoTo GotTAmt
  End If

  'Break 2
  If LastTblCnt >= 3 Then
    If MeterConsump& > RateTbl.TblBreaks(2).UNITS And MeterConsump& <= RateTbl.TblBreaks(3).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GoTo GotTAmt
  End If

  'Break 3
  If LastTblCnt >= 4 Then
    If MeterConsump& >= RateTbl.TblBreaks(3).UNITS And MeterConsump& <= RateTbl.TblBreaks(4).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GoTo GotTAmt
  End If

  'Break 4
 If LastTblCnt >= 5 Then
   If MeterConsump& >= RateTbl.TblBreaks(4).UNITS And MeterConsump& <= RateTbl.TblBreaks(5).UNITS Then
     UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
     GoTo GotTAmt
   Else
     UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
   End If
 Else
   UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
   TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
   GoTo GotTAmt
 End If

 'break 5
 If LastTblCnt >= 6 Then
   If MeterConsump& >= RateTbl.TblBreaks(5).UNITS And MeterConsump& <= RateTbl.TblBreaks(6).UNITS Then
     UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
     GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GoTo GotTAmt
  End If

  'break 6
  If LastTblCnt >= 7 Then
    If MeterConsump& >= RateTbl.TblBreaks(6).UNITS And MeterConsump& <= RateTbl.TblBreaks(7).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GoTo GotTAmt
  End If

  'break 7
  If LastTblCnt >= 8 Then
    If MeterConsump& >= RateTbl.TblBreaks(7).UNITS And MeterConsump& <= RateTbl.TblBreaks(8).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GoTo GotTAmt
  End If
  'break 8
  If LastTblCnt >= 9 Then
    If MeterConsump& >= RateTbl.TblBreaks(8).UNITS And MeterConsump& <= RateTbl.TblBreaks(9).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GoTo GotTAmt
  End If

  'break 9
  If LastTblCnt >= 10 Then
    If MeterConsump& >= RateTbl.TblBreaks(9).UNITS And MeterConsump& <= RateTbl.TblBreaks(10).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GoTo GotTAmt
  End If

GotTAmt:
  GetRevCharge# = Round#(MinBillAmt# + TAmt#)

End Function

Public Function Check99File()
  Dim zz As Integer, Ext As String, NewName As String
  If Exist("UBPAY99.DAT") Then
    UBLog "DRAFT: UBPAY99.DAT ALLREADY EXISTS!"
    
    If MsgBox("Draft payment file already exits. Delete or Cancel?", vbOKCancel, "Continue?") = vbOK Then
      Check99File = True
      GoSub RenameOld99File
    Else
      Check99File = False
    End If
  Else
    Check99File = True
  End If

Exit Function
RenameOld99File:
  For zz = 1 To 999
    Ext$ = "000" + QPTrim$(Str$(zz))
    Ext$ = Right$(Ext$, 3)
    NewName$ = "UBPAY99." + Ext$
    If Not Exist(NewName$) Then
      UBLog "DRAFT: RENAMED UBPAY99.DAT TO " + "UBPAY99." + Ext$
      Name "UBPAY99.DAT" As NewName$
      UBLog "DRAFT: PAYMENT FILE RENAMED SUCCESSFULLY"
      Exit For
    End If
  Next
Return
End Function
Public Sub KillACHFiles()
  Dim KCnt As Integer, FileCount As Integer, FileName As String
  Dim cnt As Integer
  ReDim FileSpec$(2)
  FileSpec$(1) = "DS*."
  FileSpec$(2) = "*.ACH"

  For KCnt = 1 To 2
    ReDim TempArray$(0)
    FileCount = 0

    FileName$ = Dir$(FileSpec$(KCnt))
    If Len(FileName$) = 0 Then
      GoTo ExitKill
    Else
      FileCount = 1                    'It is, so count files.
      ReDim Preserve TempArray$(FileCount)
      TempArray$(FileCount) = FileName$
      Do
        FileName$ = Dir$
        If Len(FileName$) = 0 Then
          Exit Do
        Else
          FileCount = FileCount + 1
          ReDim Preserve TempArray$(FileCount)
          TempArray$(FileCount) = FileName$
        End If
      Loop
    End If
    For cnt = 1 To FileCount
      KillFile TempArray$(cnt)
    Next
ExitKill:
  Next
End Sub
Public Sub FixBrokenCustFile(xxcust As Long)
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim CustFile As Integer, cnt As Long, CustFile2 As Integer
 ' ReDim UBCustRecT(1) As NewUBCustRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  CustFile2 = FreeFile
  Open UBPath$ + "UBCUSTTT.DAT" For Random Shared As CustFile2 Len = UBCustRecLen
  Dim nname As String
  For cnt = 1 To xxcust
    Get CustFile, cnt, UBCustRec(1)
'    If cnt = 2772 Then Stop
'    nname = UBCustRec(1).CustName
    
    Put CustFile2, , UBCustRec(1)
  Next
  Close
  UBLog "Wrote Accounts 1 to " + CStr(xxcust)
  MsgBox "This is done.", vbOKOnly
End Sub
Public Sub DeleteBlankCusts()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim CustFile As Integer, cnt As Long
 ' ReDim UBCustRecT(1) As NewUBCustRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
 
  For cnt = 2588 To 8351
    Get CustFile, cnt, UBCustRec(1)
      UBCustRec(1).Status = ""
      UBCustRec(1).Book = ""
      UBCustRec(1).SEQNUMB = ""
      UBCustRec(1).DelFlag = True
    Put CustFile, cnt, UBCustRec(1)
  Next
  Close
  UBLog "Set Accounts 2588 to 8351 to deleted"
  MsgBox "This is done.", vbOKOnly
End Sub

Public Sub FixBrokenAverage()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  Get UBCust, 1675, UBCustRec(1)
        
      For cnt = 1 To 7
            UBCustRec(1).LocMeters(cnt).UseCnt = 1
            UBCustRec(1).LocMeters(cnt).AvgUse = 0
      Next
      Put UBCust, 1675, UBCustRec(1)
  Close
  MsgBox "This is done", vbOKOnly, "Done"
  
End Sub
Public Sub FixBrokenMsgNum()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  ReDim MsgRec(1) As UBMessRecType
  Dim MsgLen As Integer
  Dim UBFile As Integer, zz As Integer
  Dim NumMsgRec As Long, MRec As Long
  

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCustRecs& = LOF(UBCust) \ UBCustRecLen
  
  
  MsgLen = Len(MsgRec(1))
  NumMsgRec& = FileSize&("UBMESAGE.DAT") / MsgLen


  
  For cnt = 1 To NumOfCustRecs&
    Get UBCust, cnt, UBCustRec(1)
     MRec& = UBCustRec(1).MessageRec
     If MRec& > 0 And MRec& > NumMsgRec& Then
       UBCustRec(1).MessageRec = 0
       Put UBCust, cnt, UBCustRec(1)
     End If
   Next
  Close
  MsgBox "This is done", vbOKOnly, "Done"
  
End Sub

Public Sub FixReadsonTrans() 'remove 3 zeros from all trans readings
  Dim UBTranRecLen As Integer, read As String
  Dim UBFile As Integer, cntm As Integer
  Dim TNumOfRecs As Long, cnt As Long, TrTyp As Integer
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  TrTyp = TranUtilityBill

  FrmShowPctComp.Label1 = "Searching Transactions"
  FrmShowPctComp.Show

    UBFile = FreeFile
    Open "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    TNumOfRecs& = LOF(UBFile) / UBTranRecLen
    For cnt& = 1 To TNumOfRecs&
      FrmShowPctComp.ShowPctComp cnt&, TNumOfRecs&
      Get UBFile, cnt&, UBTranRec(1)
      If (UBTranRec(1).TransType = TrTyp) Or (UBTranRec(1).TransType = TrTyp + 100) Then
        For cntm = 1 To 7
          read$ = Str$(UBTranRec(1).CurRead(cntm))
          If Right$(read$, 3) = "000" Then
            UBTranRec(1).CurRead(cntm) = UBTranRec(1).CurRead(cntm) / 1000
          End If
          read$ = Str$(UBTranRec(1).PrevRead(cntm))
          If Right$(read$, 3) = "000" Then
            UBTranRec(1).PrevRead(cntm) = UBTranRec(1).PrevRead(cntm) / 1000
          End If
        Next
        Put UBFile, cnt&, UBTranRec(1)
      End If
    Next
  Erase UBTranRec
  Close
  MsgBox "All finished", vbOKOnly, "Done"

End Sub
Public Sub FixCurrPrevMult()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long, read As String
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer, cntm As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCustRecs& = LOF(UBCust) \ UBCustRecLen
  For cnt = 1 To NumOfCustRecs&
  
    Get UBCust, cnt, UBCustRec(1)
      For cntm = 1 To 7
        read$ = UBCustRec(1).LocMeters(cntm).CurRead
        If Right$(read$, 3) = "000" Then
          UBCustRec(1).LocMeters(cntm).CurRead = UBCustRec(1).LocMeters(cntm).CurRead / 1000
        End If
        read$ = UBCustRec(1).LocMeters(cntm).PrevRead
        If Right$(read$, 3) = "000" Then
          UBCustRec(1).LocMeters(cntm).PrevRead = UBCustRec(1).LocMeters(cntm).PrevRead / 1000
        End If
        UBCustRec(1).LocMeters(cntm).MTRMulti = 1000
      Next
      Put UBCust, cnt, UBCustRec(1)
    Next
  Close
  MsgBox "Cust Mult fixed", vbOKOnly, "Done"
  
End Sub
Public Sub ClearMonthAmts()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long, read As String
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer, cntm As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCustRecs& = LOF(UBCust) \ UBCustRecLen
  FrmShowPctComp.Label1 = "Clearing Monthly Fields"
  FrmShowPctComp.Show

  For cnt = 1 To NumOfCustRecs&
    FrmShowPctComp.ShowPctComp cnt&, NumOfCustRecs&
    Get UBCust, cnt, UBCustRec(1)
    For cntm = 1 To 2
      UBCustRec(1).Monthly(cntm).AMTOWED = 0
      UBCustRec(1).Monthly(cntm).TotAmtPD = 0
      UBCustRec(1).Monthly(cntm).PayAmt = 0
      UBCustRec(1).Monthly(cntm).RevSource = 0
    Next
    UBCustRec(1).MFEE1 = 0
    UBCustRec(1).MFEE2 = 0
    Put UBCust, cnt, UBCustRec(1)
  Next
  Close
  UBLog "Cleared Month Amts per utils"
  MsgBox "Cust Monthly Charges fixed", vbOKOnly, "Done"
  
End Sub
Public Sub ClearBalances()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long, read As String
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer, cntm As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCustRecs& = LOF(UBCust) \ UBCustRecLen
  FrmShowPctComp.Label1 = "Clearing Balances"
  FrmShowPctComp.Show

  For cnt = 1 To NumOfCustRecs&
    FrmShowPctComp.ShowPctComp cnt&, NumOfCustRecs&
    Get UBCust, cnt, UBCustRec(1)
    For cntm = 1 To 15
      UBCustRec(1).CurrRevAmts(cntm) = 0
      UBCustRec(1).PrevRevAmts(cntm) = 0
    Next
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
   ' UBCustRec(1).DepositAmt = 0
    UBCustRec(1).LastTrans = 0
    Put UBCust, cnt, UBCustRec(1)
  Next
  Close
  UBLog "Cust balances cleared out"
  MsgBox "Cust balances cleared", vbOKOnly, "Done"
End Sub
Public Sub SetAllowPenaltyY()
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long, read As String
  Dim UBCust As Integer, cnt As Long, CustFile2 As Integer, cntm As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  NumOfCustRecs& = LOF(UBCust) \ UBCustRecLen
  FrmShowPctComp.Label1 = "Clearing Monthly Fields"
  FrmShowPctComp.Show

  For cnt = 1 To NumOfCustRecs&
    FrmShowPctComp.ShowPctComp cnt&, NumOfCustRecs&
    Get UBCust, cnt, UBCustRec(1)
       UBCustRec(1).CUTOFFYN = "Y"
       UBCustRec(1).LATEFEE = "Y"
    Put UBCust, cnt, UBCustRec(1)
  Next
  Close
  UBLog "Cleared Month Amts per utils"
  MsgBox "Cust Monthly Charges fixed", vbOKOnly, "Done"
  
End Sub

'Public Sub DisplayCustDepositTrans(CustRec As Long)
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
'  Dim PrevTranRec As Long, FoundCnt As Integer
'  Dim UBFile As Integer, dcnt As Integer, FoundCM As Integer
'  Dim Build As String * 80
'  Dim TType As String, TDesc As String
'  Dim CurBal As Double, PreBal As Double
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  FoundCnt = 0
'  FoundCM = 0
'  frmInfo.Label1 = "Loading. . ."
'  frmInfo.Show
'  DoEvents
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  UBFile = FreeFile
'  Open UBCustFile For Random Shared As UBFile Len = UBCustRecLen
'  Get UBFile, CustRec&, UBCustRec(1)
'  Close UBFile
'
'  CurBal# = UBCustRec(1).CurrBalance
'  PreBal# = UBCustRec(1).PrevBalance
''
'Top:
''
'  UBFile = FreeFile
'  Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
'
'  PrevTranRec& = UBCustRec(1).LastTrans
'  If PrevTranRec& > 0 Then
'    Do While PrevTranRec& > 0
'      dcnt = dcnt + 1
'      Get UBFile, PrevTranRec&, UBTranRec(1)
'       If UBTranRec(1).TransType = TranDepositPayment Or UBTranRec(1).TransType = TranDepositPayment + 100 Then
'        If UBTranRec(1).VoidFlag = True Then
'         'just skip to next
'        Else
'         If UBTranRec(1).FromCMFlag = True Then
'          FoundCM = FoundCM + 1
'         Else
'          TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
'          If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
'            TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
'          Else
'            TType$ = "Deposit Payment"
'          End If
'          LSet Build = " " + Num2Date(UBTranRec(1).TransDate)
'          Mid$(Build, 20) = TType$
'          Mid$(Build, 48) = Using("#####.##", UBTranRec(1).Transamt, True)
'          Mid$(Build, 63) = Using("#####.##", UBTranRec(1).RunBalance, True)
'          Mid$(Build$, 71) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
'          frmDepListing.fpTRList.AddItem Build$
'          FoundCnt = FoundCnt + 1
'         End If
'        End If
'       End If
'      PrevTranRec& = UBTranRec(1).PrevTrans
'    Loop
'  End If
'  Close UBFile
'  'frmTRDispList.Label5.Caption = QPTrim(UBCustRec(1).CustName)
'  'frmDepListing.Label2 = "Balance: " + Using("#####.##", CurBal# + PreBal#, True)
'  'frmDepListing.Label3 = "Current:  " + Using("#####.##", CurBal#, True)
'  'frmDepListing.Label4 = "Previous:  " + Using("#####.##", PreBal#, True)
'  Unload frmInfo
'  DoEvents
'  If FoundCnt > 0 And FoundCM = 0 Then
'    frmDepListing.Show vbModal
'  ElseIf FoundCM > 0 Then
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "Deposit Payment Taken in"
'    MsgText(3) = "Cash Management."
'    MsgText(4) = ""
'    MsgText(5) = "Must be voided thru CM."
'    GetOKorNot MsgText(), True
'  Else
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "Deposit Payment Voided"
'    MsgText(3) = "Can Not Be Voided Again."
'    MsgText(4) = ""
'    MsgText(5) = ""
'    GetOKorNot MsgText(), True
'  End If
'  Erase UBTranRec, UBCustRec
'
'
'End Sub

