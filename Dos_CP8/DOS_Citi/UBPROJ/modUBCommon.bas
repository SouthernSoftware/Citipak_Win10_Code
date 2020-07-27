Attribute VB_Name = "modUBCommon"
Option Explicit

Public Sub UPDateOK()
  frmDataUpdated.Show vbModal
'  Stop
End Sub
'!!! Procedures below Needed for reports!!! Mark with!!!
'Make sure to check w/Dale  PS
'!!! Added round on 4-17-03
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function
'!!! populates the combo box with revenues
Public Function FillRevList(x As fpCombo)
  ReDim UBSetUpRec(1) As UBSetupRecType
  Dim RecLen As Integer
  Dim cnt As Integer
  LoadUBSetUpFile UBSetUpRec(), RecLen
  x.AddItem "All Revenues"
  For cnt = 1 To 15
  If Trim(UBSetUpRec(1).Revenues(cnt).REVNAME) = "" Then
    Exit For
  End If
  x.AddItem Trim(UBSetUpRec(1).Revenues(cnt).REVNAME)
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
  Dim UBSetupLen As Integer, NumOfRevs As Integer, Handle As Integer
  Dim RevCnt As Integer, TempRev As String
  NumOfRevs = 15
  ReDim UBSetUp(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUp(1))
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  Handle = FreeFile
'  Open UBPath$ + "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen    'open data file
'  Get #Handle, 1, UBSetUpRec(1)
    LoadUBSetUpFile UBSetUp(), UBSetupLen

  For RevCnt = 1 To 15
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(RevCnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    End If
  Next
  GetNumOfRevs = NumOfRevs
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
    SequenceIndex(cnt).SeqNumber = UBCustRec(1).SEQ
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
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

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
'Remark all rest of code but wanted to keep from old use the
'zip sub below Per Dale.
  MakeMowZipCodeIndex IndexText$
  Exit Sub

'  ShowProcessingScrn "Creating " + IndexText$ + " Index "
'  QPrintRC "    Reading Customer Records     ", 11, 25, -1
'
'  'REDIM ZipIndex(1 TO 1)  AS PSAZipIndexType
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
'
'  CHandle = FreeFile
'  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'
'  'REDIM ZipIndex(1 TO NumOfBillRec)   AS PSAZipIndexType
'
'  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
'    For BCnt = 1 To NumOfBillRec
'    Get CHandle, BCnt, UBCustRec(1)
'    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    'ZipIndex(BCnt).SName = UBCustRec(1).SEARCH
'    ZipIndex(BCnt).RecNum = BCnt
'    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
'  Next
'
'  Close
'
'  QPrintRC "         Sorting Index.        ", 11, 25, -1
'
'  SortT ZipIndex(1), NumOfBillRec, 0, 16, 0, 10
'
'  First = 1
'  Last = 1
'
'  SZip$ = ZipIndex(1).ZIPCODE
'
'  For ZCnt = 2 To NumOfBillRec
'    EZip$ = ZipIndex(ZCnt).ZIPCODE
'     If SZip$ <> EZip$ Then
'      Last = ZCnt - 1
'      GoSub SortThisZip
'      First = ZCnt
'      SZip$ = EZip$
'    End If
'    ShowPctComp ZCnt, NumOfBillRec              'show user percentage complete
'  Next
'  Last = ZCnt - 1
'  GoSub SortThisZip
'
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'
'  IHandle = FreeFile
'  Open TempIndexName For Output As IHandle
'  Close IHandle
'
'  IHandle = FreeFile
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For cnt = 1 To NumOfBillRec
'    Prec& = ZipIndex(cnt).RecNum
'      Put IHandle, cnt, Prec&
'    ShowPctComp cnt, NumOfBillRec               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ZipIndex

' Exit Sub

'SortThisZip:
'  If First < Last Then
'    'SortT ZipIndex(First), Last - First + 1, 0, 32, 10, 10
'    SortT ZipIndex(First), Last - First + 1, 0, 16, 10, 10
'  End If
'Return

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
Public Function FmtBook$(Book$)
  Dim BookLen As Integer
  
  Book$ = QPTrim$(Book$)
  BookLen = Len(Book$)
  
  Select Case BookLen
  Case 0
    FmtBook$ = "00"
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
  
  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
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

Public Sub UBTerminate()
  Dim UBFrmCnt As Integer
  ' Loop through the forms collection and unload each form.
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
  Dim UBFile As Integer, DCnt As Integer
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
      DCnt = DCnt + 1
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
  Case Else
    TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select

Return

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
    'QPrintRC "Has Invalid Reading!", 10, 37, -1
    MsgText(3) = "Has Invalid Reading!"

  Case 2
    'QPrintRC "Invalid Book Number!", 10, 37, -1
    MsgText(3) = "Invalid Book Number!"

  Case 3
    'AcCol = 21
    'QPrintRC "Has an INVALID RATE CODE!!", 10, 35, -1
    MsgText(3) = "Has an INVALID RATE CODE!!"

  Case 4
    'QPrintRC "Has Mismatched Meters!", 10, 37, -1
    MsgText(3) = "Has Mismatched Meters!"

  Case 5
   ' QPrintRC "Has an INVALID Reading!", 10, 37, -1
    MsgText(3) = "Has an INVALID Reading!"

  Case 6
    'QPrintRC "INVALID Flat Rate Info!", 10, 37, -1
    MsgText(3) = "INVALID Flat Rate Info!"

  Case 7
    'QPrintRC "INVALID Monthly Billed Code!", 10, 35, -1
    MsgText(3) = "INVALID Monthly Billed Code!"

  Case 8
    'QPrintRC "Meters with NO RATE Code!", 10, 36, -1
    MsgText(3) = "Meters with NO RATE Code!"

  Case 9
    'QPrintRC "Invalid Customer Type!", 10, 36, -1
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
'  QPrintRC "RATE CODE:  " + QPTrim$(RATECODE$), 10, 22, -1
'  QPrintRC "Has an INVALID entry!", 10, 39, -1
'  QPrintRC "Correct and Print Again.", 12, 28, -1
'  WaitForAction
' AbortFlag = True
'
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

