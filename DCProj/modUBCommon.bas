Attribute VB_Name = "modUBCommon"
Option Explicit
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
Public Function OKApplyDep(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBCustF As Integer
  UBCustRecLen = Len(UBCustRec(1))
  If RecNo& > 0 Then
    UBCustF = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
    Get UBCustF, RecNo&, UBCustRec(1)
    Close UBCustF
  
    If UBCustRec(1).DepositAmt <= 0 Then
      'OK = MsgBox%("UB", "NODPOSIT")
      OKApplyDep = False
    Else
      OKApplyDep = True
    End If
    If OKApplyDep = False Then
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
Public Function OKDepRefund(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBCustF As Integer
  UBCustRecLen = Len(UBCustRec(1))
  If RecNo& > 0 Then
    UBCustF = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
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

Public Function OKDepCreditAdj(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, M1 As String, M2 As String, TotalBalance As Double, UBFile As Integer
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBCustF As Integer, TNum As Long, UBTranRecLen As Integer
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  If RecNo& > 0 Then
    UBCustF = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
    Get UBCustF, RecNo&, UBCustRec(1)
    TNum = UBCustRec(1).LastTrans
    Close UBCustF
    UBFile = FreeFile
    Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    If TNum& > 0 Then
      Get UBFile, TNum&, UBTranRec(1)
    End If
    Close UBFile
    TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    If UBCustRec(1).Status = "A" Then
      M1$ = "Active Account"
      M2$ = "This Account Is Active"
      OKDepCreditAdj = False
    ElseIf TNum& <= 0 Then
      M1$ = "Missing Transactions"
      M2$ = "NO Transactions for this Account"
      OKDepCreditAdj = False
    ElseIf UBTranRec(1).TransType <> TranAppliedDeposit Then
      M1$ = "Missing Applied Deposit Transaction"
      M2$ = "Last Trans NOT An Applied Deposit"
      OKDepCreditAdj = False
    ElseIf UBCustRec(1).DepositAmt <> 0 Then
      M1$ = "Invalid Selection"
      M2$ = "This Account Has A Deposit on File"
      OKDepCreditAdj = False
    ElseIf TotalBalance# >= 0 Then
      M1$ = "Balance Not Credit"
      M2$ = "No Credit Account Balance"
      OKDepCreditAdj = False
    Else
      OKDepCreditAdj = True
    End If
    If OKDepCreditAdj = False Then
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
Public Function OKDepReverse(RecNo&)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer, M1 As String, M2 As String, TotalBalance As Double, UBFile As Integer
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  Dim dcnt As Long, FoundCM As Long, FoundCnt As Long, M3 As String
  Dim UBCustRecLen As Integer, UBCustF As Integer, PrevTranRec As Long, UBTranRecLen As Integer
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  If RecNo& > 0 Then
    UBCustF = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBCustF Len = UBCustRecLen
    Get UBCustF, RecNo&, UBCustRec(1)
    PrevTranRec& = UBCustRec(1).LastTrans
    UBFile = FreeFile
    Open UBPath + "UBTRANS.DAT" For Random Shared As UBFile Len = UBTranRecLen
    
    If PrevTranRec& <= 0 Then
      M1$ = "Missing Transactions"
      M2$ = "NO Transactions for this Account"
      OKDepReverse = False
    ElseIf UBCustRec(1).DepositAmt = 0 Then
      M1$ = "Invalid Selection"
      M2$ = "This Account Has No Deposit on File"
      OKDepReverse = False
    Else
      If PrevTranRec& > 0 Then
        Do While PrevTranRec& > 0
          dcnt = dcnt + 1
          Get UBFile, PrevTranRec&, UBTranRec(1)
           If UBTranRec(1).TransType = TranDepositPayment Or UBTranRec(1).TransType = TranDepositPayment + 100 Then
            If UBTranRec(1).VoidFlag = True Then
             'just skip to next
            Else
             If UBTranRec(1).FromCMFlag = True Then
              FoundCM = FoundCM + 1
             Else
              FoundCnt = FoundCnt + 1
             End If
            End If
           End If
          PrevTranRec& = UBTranRec(1).PrevTrans
        Loop
      End If
      If FoundCnt > 0 And FoundCM = 0 Then
       OKDepReverse = True
      ElseIf FoundCM > 0 Then
       M1$ = "Deposit Payment Taken in"
       M2$ = "Cash Management."
       M3$ = "Must be voided thru CM."
       OKDepReverse = False
     Else
      M1$ = "Deposit Payment Voided"
      M2$ = "Can Not Be Voided Again."
      OKDepReverse = False
     End If

  End If
  Close UBFile
  Erase UBTranRec, UBCustRec
  DoEvents

    If OKDepReverse = False Then
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
      MsgText(5) = M3$
      GetOKorNot MsgText(), True
    End If

  End If
End Function
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
Public Function FillGroupCMBO(x As fpCombo)
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  x.Row = 0
  x.AddItem "0" & Chr$(9) & "0" & Chr$(9) & "None"
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
    If GroupCde.Deleted = 0 Then
      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCode & Chr$(9) & GroupCde.GroupCodeName
    Else
      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCode & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
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
Public Sub LoadUBBillLetterFile(UBBillLetterRec() As UBBillLetterType, UBBillLetterlen)
  Dim Handle As Integer
  UBBillLetterlen = Len(UBBillLetterRec(1))            'use the length as an error flag
  If Exist(UBPath$ + "UBBilLtr.DAT") Then
    Handle = FreeFile
    Open UBPath$ + "UBBilLtr.DAT" For Random Shared As Handle Len = UBBillLetterlen    'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, UBBillLetterRec(1)
    End If
    Close Handle
  End If
End Sub
Public Function Exist(FileName$)
  On Local Error Resume Next
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
Public Function QPStripCom$(Address$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String
  x$ = QPTrim$(Address$)
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = "," Then
      Mid$(x$, cnt, 1) = " "
    End If
  Next

  QPStripCom$ = Trim$(x$)

End Function
Public Function QPStripLast$(NM$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String
  x$ = QPTrim$(NM$)
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, StrLen - cnt, 1)
    If thischar = " " Then
      x$ = Right$(x$, cnt)
      Exit For
    End If
  Next

  QPStripLast$ = Trim$(x$)

End Function

Public Function QPStripStuff$(Temp$)
  Dim x As String, StrLen As Long, cnt As Long, thischar As String, newcnt As Long
  Dim xx As String
  x$ = QPTrim$(Temp$)
  xx$ = ""
  newcnt = 0
  StrLen = Len(x$)
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = "(" Or thischar = ")" Or thischar = "-" Then
      Mid$(x$, cnt, 1) = " "
    End If
  Next
  For cnt = 1 To StrLen
    thischar = Mid$(x$, cnt, 1)
    If thischar = " " Then
      xx$ = xx$ + ""
    Else
      xx$ = xx$ + thischar
    End If
  Next
  QPStripStuff$ = Trim$(xx$)

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
  Dim CHandle As Integer, cnt As Long, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long, NumOfBillRec As Long
  Dim Bcnt As Long
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
  For Bcnt = 1 To NumOfBillRec
    Get CHandle, Bcnt, UBCustRec(1)
    ZipIndex(Bcnt).ZIPCODE = UBCustRec(1).ZIPCODE
    ZipIndex(Bcnt).RecNum = Bcnt
    FrmShowPctComp.ShowPctComp Bcnt, NumOfBillRec              'show user percentage complete
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
  Dim CHandle As Integer, cnt As Long, IHandle As Integer
  Dim NumCustRecs As Long, Prec As Long
  Dim Bcnt As Long

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
  If HideF7btn Then
    frmViewPrint.cmdPrnScn.Enabled = False
  End If
  DoEvents
  Unload frmLoadingRpt
  DoEvents
  frmViewPrint.Show vbModal
End Sub
Public Sub ViewPrintM(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String, Optional HideF7btn As Boolean)
 ' frmLoadingRpt.Show 1
 'not using loadingrpt form only diff between this and regular viewprint
 'the problem was all modal forms
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
 ' DoEvents
 'Unload frmLoadingRpt
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
Public Function GetDefaultLockbox%()
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  GetDefaultLockbox = Val(UBSetUpRec(1).LockBoxDef)
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
    Case "I"
      LThisMeter = MtrIrrigation
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

Public Sub UBTerminate()
  Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
  On Local Error Resume Next
  UBLog "UB Exited: "
  Ready4others PWcnt
  If DebugMode = False Then
    Shell "CitiPak.exe", vbMaximizedFocus
  End If
  DoTheTime
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
  Print #LogFile, "UB: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub
Public Sub ResetProRates()
 
  Load frmNoOperatorsWarning
  frmNoOperatorsWarning.Label(5) = "CONTINUE WITH RESET PRORATES?"
  frmNoOperatorsWarning.Show vbModal
  
  If Not DoItFlag Then
    UBLog "ABORTED: Reset Prorate Percentages"
    GoTo ExitResetProRates
  End If
 
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim CustRecLen As Integer
  Dim UBFile As Integer, CCnt As Long
  Dim NumOfCRecs As Long
  
  FrmShowPctComp.Label1 = "Reading Customer Information."
  FrmShowPctComp.AutoClose = "no"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent
  DoEvents
  
  CustRecLen = Len(UBCustRec(1))
  'NumOfCRecs& = GetNumOfCust&
  UBLog "BEGIN: Reseting Percentages"
  
  UBFile = FreeFile
  Open UBPath + "UBCUST.DAT" For Random Shared As UBFile Len = CustRecLen
  NumOfCRecs = LOF(UBFile) \ CustRecLen
  
  For CCnt = 1 To NumOfCRecs
    Get UBFile, CCnt, UBCustRec(1)
    UBCustRec(1).ProRatePCT = 100
    Put UBFile, CCnt, UBCustRec(1)
    FrmShowPctComp.ShowPctComp CCnt, NumOfCRecs
    'ShowPctComp
  Next
  Close
  Erase UBCustRec
  Unload FrmShowPctComp
  UPDateOK

ExitResetProRates:
  
End Sub

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
  frmTRDispList.Label5.Caption = QPTrim(UBCustRec(1).CustName) & "   Acct: " & Str$(CustRec&)
  frmTRDispList.Label7 = " Deposit: " + Using("#####.##", UBCustRec(1).DepositAmt)
  frmTRDispList.Label2 = "Balance: " + Using("#####.##", CurBal# + PreBal#, True)
  frmTRDispList.Label3 = "Current: " + Using("#####.##", CurBal#, True)
  frmTRDispList.Label4 = "Previous: " + Using("#####.##", PreBal#, True)
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

Public Sub RateCodeErrScrn(Ratecode$)
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "RATE CODE:  " + QPTrim$(Ratecode$)
    MsgText(3) = "Has an INVALID entry!"
    MsgText(4) = ""
    MsgText(5) = "Correct and Print Again."
    GetOKorNot MsgText(), True
End Sub

Public Static Function GetRevCharge#(RateTbl As UBRateTblRecType, TMeterConsp&, MeterMulti&)
  Dim MinBillAmt As Double, TAmt As Double, LastTblCnt As Integer
  Dim Bcnt As Integer, MeterConsump As Long, UNITS As Long
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
  For Bcnt = 1 To 10
    If RateTbl.TblBreaks(Bcnt).UNITAMT <= 0 Then
      LastTblCnt = Bcnt - 1
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

''  'break 9
''  If LastTblCnt >= 10 Then
''    If MeterConsump& >= RateTbl.TblBreaks(9).UNITS And MeterConsump& <= RateTbl.TblBreaks(10).UNITS Then
''      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
''      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
''      GoTo GotTAmt
''    Else
''      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
''      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
''    End If
''  Else
''    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
''    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
''    GoTo GotTAmt
''  End If

  'break 9
  If LastTblCnt >= 10 Then
    If MeterConsump& >= RateTbl.TblBreaks(9).UNITS And MeterConsump& <= RateTbl.TblBreaks(10).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GoTo GotTAmt
    ElseIf MeterConsump& < RateTbl.TblBreaks(9).UNITS Then
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    ElseIf MeterConsump& > RateTbl.TblBreaks(10).UNITS Then
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(10).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(10).UNITAMT))
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
Public Sub DisplayCustDepositTrans(CustRec As Long)
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim PrevTranRec As Long, FoundCnt As Integer
  Dim UBFile As Integer, dcnt As Integer, FoundCM As Integer
  Dim Build As String * 80
  Dim TType As String, TDesc As String
  Dim CurBal As Double, PreBal As Double
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  FoundCnt = 0
  FoundCM = 0
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
       If UBTranRec(1).TransType = TranDepositPayment Or UBTranRec(1).TransType = TranDepositPayment + 100 Then
        If UBTranRec(1).VoidFlag = True Then
         'just skip to next
        Else
         If UBTranRec(1).FromCMFlag = True Then
          FoundCM = FoundCM + 1
         Else
          TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
          If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
            TType$ = "Deposit Payment " + Left$(QPTrim$(UBTranRec(1).TransDesc), 10)
          Else
            TType$ = "Deposit Payment"
          End If
          LSet Build = " " + Num2Date(UBTranRec(1).TransDate)
          Mid$(Build, 20) = TType$
          Mid$(Build, 48) = Using("#####.##", UBTranRec(1).Transamt, True)
          Mid$(Build, 63) = Using("#####.##", UBTranRec(1).RunBalance, True)
          Mid$(Build$, 71) = Chr9$ + Str$(PrevTranRec&) 'Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
          frmDepListing.fpTRList.AddItem Build$
          FoundCnt = FoundCnt + 1
         End If
        End If
       End If
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop
  End If
  Close UBFile
  'frmTRDispList.Label5.Caption = QPTrim(UBCustRec(1).CustName)
  'frmDepListing.Label2 = "Balance: " + Using("#####.##", CurBal# + PreBal#, True)
  'frmDepListing.Label3 = "Current:  " + Using("#####.##", CurBal#, True)
  'frmDepListing.Label4 = "Previous:  " + Using("#####.##", PreBal#, True)
  Unload frmInfo
  DoEvents
  If FoundCnt > 0 And FoundCM = 0 Then
    frmDepListing.Show vbModal
  ElseIf FoundCM > 0 Then
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "Deposit Payment Taken in"
    MsgText(3) = "Cash Management."
    MsgText(4) = ""
    MsgText(5) = "Must be voided thru CM."
    GetOKorNot MsgText(), True
  Else
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = "Deposit Payment Voided"
    MsgText(3) = "Can Not Be Voided Again."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If
  Erase UBTranRec, UBCustRec


End Sub

Public Sub SmallPause()
Dim ST As Single
Dim ts As Single

ST = Timer + 0.2

Do
  ts = Timer
Loop Until ts >= ST


End Sub
 
 Public Sub PrintCustInfo(Rec As Long, RptType As Integer)
  Dim PageNo As Integer, Title As String, tb As Integer
  Dim Dash80 As String, ReportFile As String
  Dim UBRpt As Integer, ToPrint As String, TPDate As String
  Dim Msgflag As Boolean, RecNo As Long, NumOfRates As Integer, cnt As Integer
  Dim tmpCustRec As NewUBCustRecType
  Dim UBHandle As Integer, CustRecLen As Integer
  Title$ = "Customer Information Report"
  Dash80$ = String$(80, "-")
  TPDate$ = ""
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
      GCode$ = QPTrim$(GroupCde.GroupCode) + " " + QPTrim$(GroupCde.GroupCodeName)
    Else
      GCode$ = QPTrim$(GroupCde.GroupCode) + " Inactive"
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
  ToPrint$ = Str$(RecNo) + "~" + tmpCustRec.Book + "~" + tmpCustRec.SEQNUMB
  ToPrint$ = ToPrint$ + "~" + tmpCustRec.Status + "~" + Num2Date(tmpCustRec.OPENDATE)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.SEARCH)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.CustName)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.ADDR1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.ADDR2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.ServAddr)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.CITY)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.STATE)
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
  ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.ProRatePCT))
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG1)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG2)
  ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.HHMSG3)

  For cnt = 0 To 14
    ToPrint$ = ToPrint$ + "~" + QPTrim$(UBSetUpRec(1).Revenues(cnt + 1).RevName)
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.serv(cnt + 1).Ratecode)
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
    ToPrint$ = ToPrint$ + "~" + QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).AMTOWED))
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
    ToPrint$ = ToPrint$ + "~" + QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrUnit)
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
  ARptCustInfo.txtDate = Now
  ARptCustInfo.txtTown = TOWNNAME$
  ARptCustInfo.GetName ReportFile$
  ARptCustInfo.startrpt
  Else
  Print #UBRpt, Chr$(27); Chr$(48); Chr$(27); Chr$(58);
  Print #UBRpt, Tab(30); Title$
  Print #UBRpt, Now
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, Dash80$
  Print #UBRpt,
  Print #UBRpt, "Customer Name: "; QPTrim$(tmpCustRec.CustName);
  Print #UBRpt, Tab(46); "Status: "; tmpCustRec.Status
  Print #UBRpt, "Account #: "; Str$(RecNo); Tab(25); "Location: "; tmpCustRec.Book; "-"; tmpCustRec.SEQNUMB;
  Print #UBRpt, Tab(50); "Account Opened: "; Num2Date(tmpCustRec.OPENDATE)
  Print #UBRpt, "Address: "; QPTrim$(tmpCustRec.ADDR1);
  Print #UBRpt, Tab(50); "Group Code:  " & GCode$
  Print #UBRpt, Tab(9); QPTrim$(tmpCustRec.ADDR2);
  Print #UBRpt, Tab(46); "----Account Balance Information----"
  Print #UBRpt, Tab(9); QPTrim$(tmpCustRec.CITY); " "; QPTrim$(tmpCustRec.STATE); " "; QPTrim$(tmpCustRec.ZIPCODE);
  Print #UBRpt, Tab(50); "Account Balance: "; Using$("$###,###,###.##", Str$((Round#(tmpCustRec.CurrBalance + tmpCustRec.PrevBalance))))
  Print #UBRpt, "Service Address: "; QPTrim$(tmpCustRec.ServAddr);
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
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(1).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(1).RMtrType)
  Print #UBRpt, "Zone: "; QPTrim$(tmpCustRec.ZONE);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(2).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(2).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(2).RMtrType)
  Print #UBRpt, "Read Seq: "; QPTrim$(Str$(tmpCustRec.Seq));
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(3).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(3).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(3).RMtrType)

  Select Case tmpCustRec.CASHONLY
  Case "N", " "
    Print #UBRpt, "Cash Only: "; "No";
  Case Else
    Print #UBRpt, "Cash Only: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(4).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(4).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(4).RMtrType)
  Select Case tmpCustRec.LATEFEE
  Case "N", " "
    Print #UBRpt, "Late Fee: "; "No";
  Case Else
    Print #UBRpt, "Late Fee: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(5).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(5).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(5).RMtrType)

  Select Case tmpCustRec.CUTOFFYN
  Case "N", " "
    Print #UBRpt, "Allow Cutoff: "; "No";
  Case Else
    Print #UBRpt, "Allow Cutoff: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(6).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(6).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(6).RMtrType)
  Select Case tmpCustRec.TAXEXPT
  Case "N", " "
    Print #UBRpt, "Tax Exempt: "; "No";
  Case Else
    Print #UBRpt, "Tax Exempt: "; "Yes";
  End Select
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(7).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(7).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(7).RMtrType)

  Print #UBRpt, "Senior Citizen: "; QPTrim$(tmpCustRec.SRCIT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(8).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(8).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(8).RMtrType)
  
  Print #UBRpt, "Bill Comment: "; QPTrim$(tmpCustRec.BILLCMNT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(9).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(9).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(9).RMtrType)

  Print #UBRpt, "Pay Comment: "; QPTrim$(tmpCustRec.PAYCMNT);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(10).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(10).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(10).RMtrType)

  Print #UBRpt, "Pump Code: "; QPTrim$(tmpCustRec.PumpCode);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(11).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(11).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(11).RMtrType)

  Print #UBRpt, "User Code 1: "; QPTrim$(tmpCustRec.USERCODE1);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(12).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(12).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(12).RMtrType)

  Print #UBRpt, "User Code 2: "; QPTrim$(tmpCustRec.USERCODE2);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(13).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(13).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(13).RMtrType)

  Print #UBRpt, "Prorate%: "; QPTrim$(Str$(tmpCustRec.ProRatePCT));
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(14).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(14).Ratecode);
  Print #UBRpt, Tab(75); QPTrim$(tmpCustRec.serv(14).RMtrType)

  Print #UBRpt, "HH Message 1: "; QPTrim$(tmpCustRec.HHMSG1);
  Print #UBRpt, Tab(39); Left$(UBSetUpRec(1).Revenues(15).RevName, 18);
  Print #UBRpt, Tab(60); QPTrim$(tmpCustRec.serv(15).Ratecode);
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
    Print #UBRpt, Tab(48); QPTrim$(Str$(tmpCustRec.Monthly(cnt + 1).AMTOWED));
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
    Print #UBRpt, Tab(18); QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrUnit);
    Print #UBRpt, Tab(20); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).NumUser));
    TPDate$ = Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate)
    Print #UBRpt, Tab(22); TPDate$; 'Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate);
    If tmpCustRec.LocMeters(cnt + 1).CurRead > 0 Then
      Print #UBRpt, Tab(33); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).CurRead));
    End If
    If tmpCustRec.LocMeters(cnt + 1).PrevRead > 0 Then
      Print #UBRpt, Tab(41); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).PrevRead));
    End If
    TPDate$ = Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate)
    Print #UBRpt, Tab(49); TPDate$; 'Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate);
    TPDate$ = Num2Date(tmpCustRec.LocMeters(cnt + 1).PastDate)
    Print #UBRpt, Tab(60); TPDate$; 'Num2Date(tmpCustRec.LocMeters(cnt + 1).PastDate);
    Print #UBRpt, Tab(71); QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrIDNO);
    Print #UBRpt, Tab(77); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).MtrLat));
    Print #UBRpt, Tab(83); QPTrim$(Str$(tmpCustRec.LocMeters(cnt + 1).MtrLng))
  Next
  
  Print #UBRpt,
  Print #UBRpt, Dash80$
  Print #UBRpt, Chr$(12)

  Close

  ViewPrint ReportFile$, Title$
  KillFile ReportFile$
  End If
End Sub

Public Function GetRPTName(Newrp As String)
  Dim Part As Double
  Part = Timer
  Newrp = Newrp + QPTrim(Str(CLng(Part)))
End Function

Public Function GCodesList(x As fpList)
  Dim GroupCde As GroupCodeRecType
  Dim GrpCodeRecLen As Integer, ghandle As Integer, cnt As Integer
  Dim NumofGrps As Integer
  GrpCodeRecLen = Len(GroupCde)
  
  ghandle = FreeFile
  Open UBPath$ + "UBGrpCde.DAT" For Random Shared As ghandle Len = GrpCodeRecLen
  NumofGrps = LOF(ghandle) \ GrpCodeRecLen
  x.Row = 0
  For cnt = 1 To NumofGrps
    Get #ghandle, cnt, GroupCde
    If GroupCde.Deleted = 0 Then
      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCode & Chr$(9) & GroupCde.GroupCodeName
'    Else
'      x.AddItem Str$(cnt) & Chr$(9) & GroupCde.GroupCODE & Chr$(9) & "Inactivated Code"
    End If
  Next
  Close
End Function

Public Sub GetGCodestoReport(x As fpList)
  Dim CodestoRpt As GroupCodeRptType
  Dim Codefile As Integer
  Dim CodeName As String
  Dim PCnt As Integer, cnt As Integer
  '--process  only the selected choices
  CodeName$ = "grpcds.LST"
  KillFile CodeName$
  Codefile = FreeFile
  Open CodeName$ For Random As Codefile Len = Len(CodestoRpt)
  x.ListIndex = 0
  'If fplstFunds.Selected = True Then
    For PCnt = 0 To x.ListCount - 1
      If x.Selected(PCnt) Then
        cnt = cnt + 1
        x.col = 0
        x.ListIndex = PCnt
        CodestoRpt.RecordNum = QPTrim$(x.ColText)
        x.col = 1
        CodestoRpt.GroupCode = QPTrim$(x.ColText)
        Put Codefile, cnt, CodestoRpt
        
      End If
    Next
  Close
End Sub
Public Sub GetRCodestoReport(x As fpList)
  Dim RCstoRpt As RateCodeRptType
  Dim RCfile As Integer
  Dim RCName As String
  Dim PCnt As Integer, cnt As Integer
  '--process  only the selected choices
  RCName$ = "Ratecds.LST"
  KillFile RCName$
  RCfile = FreeFile
  Open RCName$ For Random As RCfile Len = Len(RCstoRpt)
  x.ListIndex = 0
  'If fplstFunds.Selected = True Then
    For PCnt = 1 To x.ListCount - 1
      If x.Selected(PCnt) Then
        cnt = cnt + 1
        x.col = 0
        x.ListIndex = PCnt
        RCstoRpt.RecordNum = PCnt
        x.col = 0
        RCstoRpt.Ratecode = QPTrim$(x.ColText)
        Put RCfile, cnt, RCstoRpt
        
      End If
    Next
  Close
End Sub
Public Sub PrnOneWO(Custlook&, WOlook&)
  Unload frmRptWrkOrdHist
  frmReportOpt.Show 1
  PRintOneWO Custlook&, WOlook&, rptopt
End Sub

Public Sub PRintOneWO(Custlook&, WOlook&, RptType As Integer)
  Dim TDate As String, cnt As Integer, TransRecNum As Long
  Dim UBWrkOrd As Integer, WorkOrderRecLen As Integer
  Dim Header As String, MtrCnt As Integer, Acct As Long
  Dim Rem1 As String, Rem2 As String, Rem3 As String, Rem4 As String
  Dim Rem5 As String, Rem6 As String, ToPrint As String
  Dim dcnt As Long, PCnt As Integer, Dash As String
  Dim Rec As Long, CustNum As Long
  Dim UBCustRecLen As Integer, UBFile As Integer
  Dim ReportFile As String, RptHandle As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  CustNum = Custlook
  TransRecNum& = WOlook
  ReDim WorkOrderRec(1) As WorkOrderRecType
  WorkOrderRecLen = Len(WorkOrderRec(1))
  Header$ = "Work Order"
  UBWrkOrd = FreeFile
  Open UBPath$ + "UBWRKORD.DAT" For Random Shared As UBWrkOrd Len = WorkOrderRecLen
  If TransRecNum& > 0 Then
    Get UBWrkOrd, TransRecNum&, WorkOrderRec(1)
  End If
  Close UBWrkOrd
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  Acct = CustNum
  UBFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustNum&, UBCustRec(1)
  Close UBFile
  ReportFile$ = UBPath$ + "WORKORDR.RPT"   'Open Report File
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  Rem1$ = ""
  Rem2$ = ""
  Rem3$ = ""
  Rem4$ = ""
  Rem5$ = ""
  Rem6$ = ""
  If rptopt = 0 Then
    Dash$ = String$(79, "_")
  Else
    Dash$ = String$(83, "_")
  End If

If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(1))) > 0 Then
    Rem1$ = QPTrim(WorkOrderRec(1).RepliesText.Text(1))
  Else
    Rem1$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(2))) > 0 Then
    Rem2$ = QPTrim(WorkOrderRec(1).RepliesText.Text(2))
  Else
    Rem2$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(3))) > 0 Then
    Rem3$ = QPTrim(WorkOrderRec(1).RepliesText.Text(3))
  Else
    Rem3$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(4))) > 0 Then
    Rem4$ = QPTrim(WorkOrderRec(1).RepliesText.Text(4))
  Else
    Rem4$ = Dash$
  End If
  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(5))) > 0 Then
    Rem5$ = QPTrim(WorkOrderRec(1).RepliesText.Text(5))
  Else
    Rem5$ = Dash$
  End If

  If Len(QPTrim(WorkOrderRec(1).RepliesText.Text(6))) > 0 Then
    Rem6$ = QPTrim(WorkOrderRec(1).RepliesText.Text(6))
  Else
    Rem6$ = "BY: ______________________________   DATE: ____________________"
  End If
'WorkOrderRec(1).CompletedDate
  If rptopt <> 1 Then
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, Tab(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "    Work Order#: "; Using("######", UBCustRec(1).WOLastTrans); Tab(30); "Date Issued: "; Num2Date$(WorkOrderRec(1).ENTRYDATE)
    Print #RptHandle, "      Location#: "; UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; Tab(30); "Complete By: "; Num2Date$(WorkOrderRec(1).CompleteByDate)
    Print #RptHandle, "       Account#: "; Acct&; Tab(30); "  Completed: "; Num2Date$(WorkOrderRec(1).CompletedDate)
    Print #RptHandle, "  Customer Name: "; UBCustRec(1).CustName
    Print #RptHandle, "Service Address: "; UBCustRec(1).ServAddr
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Instruction or Description of Work Needed"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(1)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(2)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(3)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(4)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(5)
    Print #RptHandle, WorkOrderRec(1).OrdersText.Text(6)
    Print #RptHandle, " "
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, "Remarks Noted by Worker"
    Print #RptHandle, String$(79, "-")
    Print #RptHandle, Rem1$
    Print #RptHandle, " "
    Print #RptHandle, Rem2$
    Print #RptHandle, " "
    Print #RptHandle, Rem3$
    Print #RptHandle, " "
    Print #RptHandle, Rem4$
    Print #RptHandle, " "
    Print #RptHandle, Rem5$
    Print #RptHandle, " "
    Print #RptHandle, Rem6$
    Print #RptHandle, " "
    Print #RptHandle, " "
    Print #RptHandle, "Meter Numbers:"

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #RptHandle, QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      End If
    Next
    Print #RptHandle, FF$;
  Else
    ToPrint$ = Num2Date$(WorkOrderRec(1).ENTRYDATE) + "~"
    ToPrint$ = ToPrint$ + Using("######", UBCustRec(1).WOLastTrans) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB + "~"
    ToPrint$ = ToPrint$ + Str(Acct&) + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).CustName + "~"
    ToPrint$ = ToPrint$ + UBCustRec(1).ServAddr + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(1) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(2) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(3) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(4) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(5) + "~"
    ToPrint$ = ToPrint$ + WorkOrderRec(1).OrdersText.Text(6) + "~"
    ToPrint$ = ToPrint$ + Rem1$ + "~"
    ToPrint$ = ToPrint$ + Rem2$ + "~"
    ToPrint$ = ToPrint$ + Rem3$ + "~"
    ToPrint$ = ToPrint$ + Rem4$ + "~"
    ToPrint$ = ToPrint$ + Rem5$ + "~"
    ToPrint$ = ToPrint$ + Rem6$

    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        ToPrint$ = ToPrint$ + "~" + QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      Else
        ToPrint$ = ToPrint$ + "~ "
      End If
    Next
    ToPrint$ = ToPrint$ + "~" + Num2Date$(WorkOrderRec(1).CompleteByDate) + "~"
    ToPrint$ = ToPrint$ + Num2Date$(WorkOrderRec(1).CompletedDate)

    Print #RptHandle, ToPrint$
    ToPrint$ = ""
  End If
  Close
  If rptopt <> 1 Then
      ViewPrint ReportFile$, Header$
    Else
      Load frmLoadingRpt
      'frmLoadingRpt.setwherefrom frmRptWrkOrdHist
      ARptWorkOrder.GetName ReportFile$
      ARptWorkOrder.startrpt
    End If

End Sub


