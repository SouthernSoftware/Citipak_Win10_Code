Attribute VB_Name = "modUBLookup"
Option Explicit



Public Function LookUp&(LookFor$, FindType%, ClearScrn%, ActiveOnly%, ParentForm As Form)

  Dim AcctNum As Long, TCnt As Integer
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBCustSN(1) As nUBCustReIndexRecType
  Dim UBCustRecLen As Integer, UBCustSNLen As Integer
  Dim C1Handle As Integer, R1Handle As Integer, dcnt As Integer
  Dim SearchLen As Integer, AbortFlag As Integer
  Dim NumOfCust As Long, CCnt As Long
  Dim UBSearchN As String, Build As String * 80
  Dim TCustName As String
  Dim OK2Search As Integer, DashPos As Integer
  Dim LNum As String, Book As String, SeqN As String
  Dim SAddrFlag As Integer, AddrOKFlag As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, MidRec As Long
  Dim FirstRec As Long, LastRec As Long, LastSRec As Long
  Dim BotOffSet As Long, TopOffSet As Long, FirstMatchRec As Long

  UBCustRecLen = Len(UBCustRec(1))

  Select Case FindType
  Case 2, 3, 4, 6:   'all but account and location lookups
    Load frmDisplayList
  Case Else:
  End Select

  LookFor$ = UCase$(LookFor$)

  Select Case FindType
  Case 1    'account lookup
    AcctNum& = Val(LookFor$)
    If AcctNum& < 1 Or AcctNum& > GetNumOfCust Then
      Load frmLookupError
      frmLookupError.Label = "Invalid Account Number!"
      frmLookupError.Show vbModal
      LookUp& = 0
    Else
      If IsDeleted(AcctNum&) Then
        Load frmLookupError
        frmLookupError.Label = "Deleted Account!"
        frmLookupError.Show vbModal
        LookUp& = 0
      Else
        LookUp& = AcctNum&
      End If
    End If
  Case 2    'Name lookup
    If Len(LookFor$) = 0 Then
      LookFor$ = Space$(10)
    End If
    GoSub Search4Cust
    If AbortFlag Then
      GoTo ExitLookUp
    End If
    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      LookUp = 0
    Else
      frmDisplayList.Caption = "Matching Accounts"
      frmDisplayList.Label2 = "Service Address"
      frmDisplayList.Show vbModal, ParentForm
      LookUp = SearchRec
    End If
  Case 3    'meter number
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If
    GoSub Search4Meter
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      LookUp = 0
    Else
      frmDisplayList.Label2 = "Meter No."
      frmDisplayList.Show vbModal, ParentForm
      LookUp = SearchRec
    End If
  Case 4    'service address
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If

    SAddrFlag = True

    GoSub Search4SAddr
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      LookUp = 0
    Else
      frmDisplayList.Label2 = "Service Address"
      frmDisplayList.Show vbModal, ParentForm
      LookUp = SearchRec
    End If
  Case 5    'Location lookup
    If AcctNum& > 0 Then
      LookUp = AcctNum&
    End If
    OK2Search = False
    LNum$ = LookFor$
    DashPos = InStr(LNum$, "-")

    If Len(LNum$) < 2 Then  'OR DashPos <= 0 THEN
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    ElseIf DashPos > 1 Then
      Book$ = FmtBook$(Left$(LNum$, DashPos - 1))
      SeqN$ = FmtSeqN$(Mid$(LNum$, DashPos + 1))
      LNum$ = Book$ + "-" + SeqN$
      OK2Search = True
    Else
      Book$ = FmtBook$(Left$(LNum$, 2))
      SeqN$ = FmtSeqN$(Mid$(LNum$, 3))
      LNum$ = Book$ + "-" + SeqN$
      OK2Search = True
    End If
    If OK2Search Then
      ParentForm.fpSearchText = LNum$
      GoSub Search4LNum
      If AcctNum& > 0 Then
        LookUp& = AcctNum&
      ElseIf AcctNum& = 0 Then
        LookUp& = 0
        frmLookupError.Label = "No Matching Location Found"
        frmLookupError.Show vbModal
        Unload frmLookupError
      End If
    End If
  Case 6   '911 Address
    If Len(LookFor$) = 0 Then
      frmLookupError.Label = "Invalid Search!"
      frmLookupError.Show vbModal
      Unload frmLookupError
      GoTo ExitLookUp
    End If
    SAddrFlag = False
    GoSub Search4SAddr
    If AbortFlag Then
      GoTo ExitLookUp
    End If

    If dcnt = 0 Then
      frmLookupError.Label = "No Matching Account Found"
      frmLookupError.Show vbModal
      Unload frmLookupError
      LookUp = 0
    Else
      frmDisplayList.Label2 = "911 Address"
      frmDisplayList.Show vbModal, ParentForm
      LookUp = SearchRec
    End If
  End Select
  GoTo ExitLookUp

'************************************************************
Search4LNum:

  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize(UBPath$ + "UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType

  FrmShowPctComp.Label1 = "Searching for Location"
  FrmShowPctComp.Show
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.ShowPctComp 1, 10

  C1Handle = FreeFile
  Open UBPath$ + "UBCUSTBK.IDX" For Random Shared As C1Handle Len = IdxRecLen
  For CCnt = 1 To IdxNumOfRecs
    Get C1Handle, CCnt, IdxBuff(CCnt)
    FrmShowPctComp.ShowPctComp CCnt, IdxNumOfRecs
  Next
  Close C1Handle

  SearchLen = Len(LookFor$)

  FirstRec = 1
  LastRec = IdxNumOfRecs

  BotOffSet = 0
  TopOffSet = IdxNumOfRecs

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen
  MidRec = (LastRec + FirstRec) \ 2

  Do
    If LastSRec = MidRec Then
      Exit Do
    End If
    LastSRec = MidRec
    Get C1Handle, IdxBuff(MidRec).RecNum, UBCustRec(1)
    UBSearchN$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    If (LNum$ = UBSearchN$) And (UBCustRec(1).DelFlag = 0) Then
      If MidRec - BotOffSet > 1 Then
        MidRec = MidRec - 1
      Else
        FirstMatchRec = MidRec
      End If
    ElseIf LNum$ < UBSearchN$ Then             'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    Else        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      If MidRec = IdxNumOfRecs + 1 Then
        Exit Do
      End If
    End If
    If TopOffSet = BotOffSet Then
      Exit Do
    End If
  Loop Until FirstMatchRec
  Close C1Handle

  If FirstMatchRec = 0 Then
    AcctNum& = 0
  Else
    AcctNum& = IdxBuff(FirstMatchRec).RecNum
  End If

  If ActiveOnly And UBCustRec(1).Status <> "A" Then
    AcctNum& = 0
  ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status <> "I") Then
    AcctNum& = 0
  End If

ExitLSearch:
  Erase UBCustRec, IdxBuff
Return

'************************************************************
Search4SAddr:
  UBCustRecLen = Len(UBCustRec(1))
  NumOfCust& = GetNumOfCust&
  If SAddrFlag Then
    FrmShowPctComp.Label1 = "Searching for Service Address"
  Else
    FrmShowPctComp.Label1 = "Searching for 911 Address"
  End If
  FrmShowPctComp.Show

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen

  dcnt = 0
  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
        GoSub CheckLoadEM2
      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        GoSub CheckLoadEM2
      End If
    End If
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
  Next
  Close C1Handle

Return

CheckLoadEM2:
  AddrOKFlag = False
  If SAddrFlag Then
    If InStr(UBCustRec(1).ServAddr, LookFor$) > 0 Then
      AddrOKFlag = True
    End If
  Else
    If InStr(UBCustRec(1).Addr911, LookFor$) > 0 Then
      AddrOKFlag = True
    End If
  End If
  If AddrOKFlag Then
    LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 30)
    If SAddrFlag Then
      Mid$(Build$, 32, 25) = Left$(QPTrim$(UBCustRec(1).ServAddr), 25)
    Else
      Mid$(Build$, 32, 25) = QPTrim$(UBCustRec(1).Addr911)
    End If
    Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    Mid$(Build$, 74) = Chr9$ + Str$(CCnt&)
    Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
    frmDisplayList.fpList1.AddItem Build$
    dcnt = dcnt + 1
  End If
Return

'*************************************************************

Search4Meter:

  UBCustRecLen = Len(UBCustRec(1))
  NumOfCust& = GetNumOfCust&

  FrmShowPctComp.Label1 = "Searching for Meter Number"
  FrmShowPctComp.Show

  C1Handle = FreeFile
  Open UBCustFile For Random Shared As C1Handle Len = UBCustRecLen

  dcnt = 0
  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustRec(1)
    If Not UBCustRec(1).DelFlag Then
      'IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
        GoSub CheckEM2
      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
        GoSub CheckEM2
      End If
    End If
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
  Next
  Close C1Handle

Return

CheckEM2:
  For TCnt = 1 To 7
    If InStr(UBCustRec(1).LocMeters(TCnt).MtrNum, LookFor$) > 0 Then
      LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 30)
      Mid$(Build$, 32, 12) = QPTrim$(UBCustRec(1).LocMeters(TCnt).MtrNum)
      Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      Mid$(Build$, 74) = Chr9$ + Str$(CCnt&)
      Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
      frmDisplayList.fpList1.AddItem Build$
      dcnt = dcnt + 1
    End If
  Next
Return

'************************************************************
Search4Cust:
  UBCustSNLen = Len(UBCustSN(1))

  FrmShowPctComp.Label1 = "Searching Customers"
  FrmShowPctComp.Show

  SearchLen = Len(LookFor$)

  C1Handle = FreeFile
  Open UBPath$ + "UBCUSTSN.DAT" For Random Shared As C1Handle Len = UBCustSNLen
  'open short name data file
  R1Handle = FreeFile
  Open UBCustFile For Random Shared As R1Handle Len = UBCustRecLen
  'open customer data file

  NumOfCust& = LOF(C1Handle) / UBCustSNLen

  For CCnt& = 1 To NumOfCust&
    Get C1Handle, CCnt&, UBCustSN(1)
      UBSearchN$ = Left$(UBCustSN(1).SearchName, SearchLen)
      If (LookFor$ = UBSearchN$) Then
        If Len(QPTrim$(UBCustSN(1).DelFlag)) Then GoTo DelSkip2
        If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustSN(1).Status = "A"))) Then
          GoSub CustLoadEM2
        ElseIf (ActiveOnly = 1) And (UBCustSN(1).Status = "I") Then
          GoSub CustLoadEM2
        End If
      End If
DelSkip2:
    'Next
    FrmShowPctComp.ShowPctComp CCnt&, NumOfCust&
    If FrmShowPctComp.Out Then
      Unload FrmShowPctComp
      AbortFlag = True
      Exit For
    End If
    'ShowPctCompL CCnt&, NumChunks&
    'ShowSearchWheel 12, 44
  Next

  Close C1Handle               'close files
  Close R1Handle

Return

CustLoadEM2:

  Get R1Handle, UBCustSN(1).RecNum, UBCustRec(1)

  dcnt = dcnt + 1
  LSet Build$ = Left$(QPTrim$(UBCustRec(1).CustName), 26)
  Mid$(Build$, 28) = Left$(QPTrim$(UBCustRec(1).ServAddr), 30)
  Mid$(Build$, 60) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
  Mid$(Build$, 71) = QPTrim$(UBCustRec(1).Status)
  Mid$(Build$, 74) = Chr9$ + Str$(UBCustSN(1).RecNum)
  frmDisplayList.fpList1.AddItem Build$

Return
'************************************************************

ExitLookUp:
End Function

