Attribute VB_Name = "modAPCommon"
Option Explicit
  Dim VendorIdx As VendorIdxRecType
  Dim Vendor As VendorRecType
  Dim POEdit As POFORMRecType2
  Dim POTrans As GLTransRecType
  Dim POCont(1) As POControlRecType
  Public New1099 As Boolean
  Public OKtoPO As Boolean
'----- Look up tables for number words
Const NumTbl$ = "123456789"
Const NumNames$ = "One  Two  ThreeFour Five Six  SevenEightNine Ten"
Const Teens$ = "Eleven    Twelve    Thirteen  Fourteen  Fifteen   Sixteen   Seventeen Eighteen  Nineteen"
Const Tens$ = "Ten     Twenty  Thirty  Forty   Fifty   Sixty   Seventy Eighty  Ninety"

Const Powers3$ = "Thousand Million  Billion  Trillion"

'DECLARE FUNCTION SpellNumber$ (Number$)
'******* Returns a spelled out version of a number
Public Function SpellNumber$(StrNum$)    ' STATIC
  Dim Num As String, x As Integer, Length As Integer
  Dim N As Integer, Temp As Integer, Word As String, Sentence As String
    SpellNumber$ = ""                           'Clear the function
    Num$ = LTrim$(RTrim$(StrNum$))              'Trim off any spaces

    x = InStr(Num$, ".")                        'Trim off any decimal places
    If x Then Num$ = Left$(Num$, x - 1)

    Length = Len(Num$)                          'Get the length
    If Length > 15 Then Exit Function           'Exit if bigger than trillions

    For N = Length To 1 Step -1                 'Step backwards through number

        x = InStr(NumTbl$, Mid$(Num$, N, 1)) - 1 'Look up the digit in table

        Select Case (Length - N) Mod 3          'Branch according to digit
                                                '  position
           '----- Ones digit
           Case 0
              If N < Length Then                'If not on last digit, look
                 For Temp = N To N - 2 Step -1  '  for non 0 digit
                    If Temp > 0 Then            'If not past end of number
                       Word$ = Mid$(Num$, Temp, 1)
                                                'If this is a non 0 digit,
                                                '  put power word in sentence
                       If Word$ <> "0" And Word$ <> "-" Then
                          Temp = ((Length - N) \ 3 - 1) * 9 + 1
                          Word$ = RTrim$(Mid$(Powers3$, Temp, 9))
                          Sentence$ = Word$ + " " + Sentence$
                          Exit For              'Bail out of search loop
                       End If
                    End If
                 Next
              End If

              If x > -1 Then                    'If digit found, get the word
                 Word$ = Mid$(NumNames$, x * 5 + 1, 5)

                 If N > 1 Then                  'If left digit is one, use
                                                '  "Teen" table
                    If Mid$(Num$, N - 1, 1) = "1" Then
                       Word$ = Mid$(Teens$, x * 10 + 1, 10)
                       N = N - 1                'Skip the Tens digit
                    End If
                 End If
              End If

           '----- Tens digit
           Case 1
              If x > -1 Then                    'Find word in "Tens" table
                 Word$ = Mid$(Tens$, x * 8 + 1, 8)
              End If

           '----- Hundreds digit
           Case 2
              If x > -1 Then                    'Find word in number table
                 Word$ = Mid$(NumNames$, x * 5 + 1, 5)
                                                'Add the word "Hundred"
                 Word$ = RTrim$(Word$) + " Hundred"
              End If

        End Select

        If N = 1 And x = -1 Then                'Look for a minus sign at
           If Mid$(Num$, N, 1) = "-" Then       '  digit one
              Word$ = "Negative"                'Add it to sentence
              x = 0
           End If
        End If
                                                'If digit is non zero, add
                                                '  the word to the sentence
        If x > -1 Then Sentence$ = RTrim$(Word$) + " " + Sentence$
    Next

'****** Added "dollars and cents" directly to spellnum
'       02/22/94

    Sentence$ = RTrim$(Sentence$)
    Sentence$ = Sentence$ + " Dollars and "

    'Sentence$ = Sentence$ + " Dollar"
    'IF INT(VAL(Num$)) <> 1 THEN
    '  Sentence$ = Sentence$ + "s and "    'Anything but "One" is plural
    'ELSE
    '  Sentence$ = Sentence$ + " and "
    'END IF

    Sentence$ = Sentence$ + Mid$(StrNum$, InStr(StrNum$, ".") + 1) + " Cents"
    'Do cents part

    SpellNumber$ = RTrim$(Sentence$)            'Assign the function

    Num$ = ""                                   'Clean up work strings
    Word$ = ""
    Sentence$ = ""

End Function

Public Function GetEncAcct(EncAcct As String)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  EncAcct = QPTrim(GLSetup.EncAcct)
  Close SetupFile
End Function
Public Function GetAPAcct(APAcct As String)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  APAcct = QPTrim(GLSetup.APAcct)
  Close SetupFile
End Function
Public Function GetAPCheck(APCheck As Integer)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  If GLSetup.APChkCode > 0 Then
    APCheck = GLSetup.APChkCode
  Else
    APCheck = 0
  End If
  Close SetupFile
End Function

Public Function FindVendorRec(VendorCode$)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim Match As Boolean, FirstRec As Integer, LastRec As Integer
  Dim LookFor As String, MiddleRec As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors

  If NumActiveVendors = 0 Then
    FindVendorRec = 0
    Close VendorIdxFile
    Exit Function
  End If

  Match = False
  FirstRec = 1
  LastRec = NumActiveVendors
  LookFor$ = QPTrim$(UCase$(VendorCode$))

  Do Until LastRec < FirstRec

    MiddleRec = (LastRec + FirstRec) \ 2

    Get VendorIdxFile, MiddleRec, VendorIdx

    If LookFor$ = QPTrim$(VendorIdx.VendorCode) Then
      Match = True
      Exit Do
    ElseIf LookFor$ < VendorIdx.VendorCode Then
      LastRec = MiddleRec - 1
    Else
      FirstRec = MiddleRec + 1
    End If

  Loop

  If Match Then
    FindVendorRec = VendorIdx.RecNum
  Else
    FindVendorRec = 0
  End If

  Close VendorIdxFile

End Function
Public Sub Q2Sort(Idxbuff() As TrSortType2, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As TrSortType2
  Dim Temp2 As TrSortType2
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).TRDATE < Temp.TRDATE
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.TRDATE < Idxbuff(lngCurHigh).TRDATE
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
      Q2Sort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      Q2Sort Idxbuff(), lngCurLow, lUBound
    End If
End Sub

Public Sub IndexVendorFile(formname As Form)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim cnt As Integer, VendorName As String, VendorNumber As String
  Dim GoodAccts As Integer, low As Integer, High As Integer
   'Delete index file if it exits
    KillFile "apvendor.idx"
  FrmShowPctComp.Label1 = "Initializing Vendor Index."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents

  '  '--Open the Index file
  OpenVendorIdx VendorIdxFile, NumActiveVendors

  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file

  If LOF(VendorFile) = 0 Then   'get out if nothing to do
    Close
    Exit Sub
  End If

  ReDim Idxbuff(1 To NumVRecs) As VendorIdxRecType              'dim the buffe

  For cnt = 1 To NumVRecs       'Load the buffer
    
    Get VendorFile, cnt, Vendor
    If Vendor.VIN > 99999 Then
      Vendor.DelFlag = -1
      Put VendorFile, cnt, Vendor
    End If

    VendorName$ = QPTrim$(Vendor.VNAME)
    If Len(VendorName$) = 0 Then
      Vendor.DelFlag = -1
      Put VendorFile, cnt, Vendor
    End If

    If Len(VendorName$) <> 0 Then
      If Asc(Left$(VendorName$, 1)) < 48 Or Asc(Left$(VendorName$, 1)) > 90 Then
        Vendor.DelFlag = -1
        Put VendorFile, cnt, Vendor
      End If
    End If

    VendorNumber$ = RTrim$(Vendor.vnum)
    If Len(VendorNumber$) = 0 Then
      Vendor.DelFlag = -1
      Put VendorFile, cnt, Vendor
    End If

    If Len(VendorNumber$) <> 0 Then
      If Asc(Left$(VendorNumber$, 1)) < 48 Or Asc(Left$(VendorNumber$, 1)) > 90 Then
        Vendor.DelFlag = -1
        Put VendorFile, cnt, Vendor
      End If
    End If

    If Vendor.DelFlag = 0 Then  'Get only active records
      FrmShowPctComp.ShowPctComp cnt, NumVRecs
      GoodAccts = GoodAccts + 1
      LSet Idxbuff(GoodAccts).VendorCode = QPTrim$(Vendor.vnum)
      'RSET IdxBuff(GoodAccts).VendorCode = QPTrim$(Vendor.VNum)
      Idxbuff(GoodAccts).RecNum = cnt
    End If
  Next cnt

  Close VendorFile              'Close the file

  '--redim with just good accts
  ReDim Preserve Idxbuff(1 To GoodAccts) As VendorIdxRecType
  FrmShowPctComp.Label1 = "Sorting Vendors...Please Wait..."
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents
  low = LBound(Idxbuff)
  High = UBound(Idxbuff)
  FrmShowPctComp.ShowPctComp 15, 100

  QAPSort Idxbuff(), low, High
  'SortT Idxbuff(), GoodAccts

  'FPutAH "apvendor.idx", IdxBuff(1), 12, GoodAccts
  FrmShowPctComp.ShowPctComp 95, 100
  '--write out to the index file
  For cnt = 1 To GoodAccts
    FrmShowPctComp.ShowPctComp cnt, GoodAccts
    'Get VendorIdxFile, cnt, AcctIdx
    'RSET VendorIdx.VendorCode = QPTrim$(IdxBuff(Cnt).VendorCode)
    LSet VendorIdx.VendorCode = QPTrim$(Idxbuff(cnt).VendorCode)
    VendorIdx.RecNum = Idxbuff(cnt).RecNum
    Put VendorIdxFile, cnt, VendorIdx
  Next
  '
  Close VendorIdxFile           'close index

End Sub

Public Sub QAPSort(Idxbuff() As VendorIdxRecType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As VendorIdxRecType
  Dim Temp2 As VendorIdxRecType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Sub 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = Idxbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While Idxbuff(lngCurLow).VendorCode < Temp.VendorCode
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.VendorCode < Idxbuff(lngCurHigh).VendorCode
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
      QAPSort Idxbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QAPSort Idxbuff(), lngCurLow, lUBound
    End If
End Sub
Public Function QCkSort(CHKbuff() As CheckInfoType3, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As CheckInfoType3
  Dim Temp2 As CheckInfoType3
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Function 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = CHKbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While CHKbuff(lngCurLow).StartChk < Temp.StartChk
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.StartChk < CHKbuff(lngCurHigh).StartChk
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = CHKbuff(lngCurLow)
        CHKbuff(lngCurLow) = CHKbuff(lngCurHigh)
        CHKbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QCkSort CHKbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QCkSort CHKbuff(), lngCurLow, lUBound
    End If
End Function
Public Function QCkSort2(CHKbuff() As GLTransRecType, lLBound, lUBound)
  Dim lngCurMid As Long, lngCurLow As Long, lngCurHigh As Long
  Dim Temp As GLTransRecType
  Dim Temp2 As GLTransRecType
  lngCurLow = lLBound
  lngCurHigh = lUBound
  'this is to exit loop if high and low are equal
  If lUBound <= lLBound Then Exit Function 'GoTo Proc_Exit
    'find the midpoint
    lngCurMid = (lUBound + lLBound) \ 2
    '
    Temp = CHKbuff(lngCurMid)
    Do While (lngCurLow <= lngCurHigh)
      Do While CHKbuff(lngCurLow).AcctNum < Temp.AcctNum
        lngCurLow = lngCurLow + 1
        If lngCurLow = lUBound Then Exit Do
      Loop
      Do While Temp.AcctNum < CHKbuff(lngCurHigh).AcctNum
        lngCurHigh = lngCurHigh - 1
        If lngCurHigh = lLBound Then Exit Do
      Loop
      'if low is <= high then swap
      If (lngCurLow <= lngCurHigh) Then
        Temp2 = CHKbuff(lngCurLow)
        CHKbuff(lngCurLow) = CHKbuff(lngCurHigh)
        CHKbuff(lngCurHigh) = Temp2
    '
        lngCurLow = lngCurLow + 1
        lngCurHigh = lngCurHigh - 1
      End If
    Loop
  'recurse if necessary
    If lLBound < lngCurHigh Then
      QCkSort2 CHKbuff(), lLBound, lngCurHigh
    End If
  'recurse if necessary
    If lngCurLow < lUBound Then
      QCkSort2 CHKbuff(), lngCurLow, lUBound
    End If
End Function

'**** Why This not used on add new vendor, only on add vendor in invoices
'*****
Function GetNewVendorPIN()
  Dim VPinHandle As Integer, VPin As Long
  VPinHandle = FreeFile
  Open "VendrPIN.Dat" For Random Shared As #VPinHandle Len = 4
  Get #VPinHandle, 1, VPin&
  VPin& = VPin& + 1
  Put #VPinHandle, 1, VPin&
  Close #VPinHandle
  GetNewVendorPIN = VPin&
End Function
'Public Static Function FillVendNumName(txtField As fpCombo)
'  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer
'  Dim AcctFile As Integer, NumAccts As Integer, CntA As Integer
'  OpenAcctFile AcctFile
'  OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
'  NumAccts = LOF(AcctFile) / Len(GLAcct)
'  txtField.Row = -1
'  For CntA = 1 To NumAIdxRecs
'    Get AcctIdxFileNum, CntA, GLAcctidx
'    Get AcctFile, GLAcctidx.RecNum, GLAcct
'      If GLAcct.Deleted = 0 Then
'        txtField.InsertRow = Str$(GLAcctidx.RecNum) & Chr$(9) & QPTrim(GLAcct.Num) & Chr$(9) & Trim(GLAcct.title) & Chr$(9) & QPStrip(GLAcct.Num)
'      End If
'  Next
'  Close AcctIdxFileNum
'  Close AcctFile
'  'Erase AcctIdxFileNum, NumAIdxRecs
'  'Erase AcctFile, NumAccts, CntA
'  End Function
Public Function VendsList(x As fpList)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file
  x.Clear
  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
      If Vendor.DelFlag = 0 Then
        x.InsertRow = Vendor.vnum & Chr$(9) & Vendor.VNAME
      End If
  Next
  Close VendorIdxFile
  Close VendorFile
End Function
Public Function VendsLstAlpha(x As fpList)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs 'Open the Acct file
  x.Clear
  x.SortState = SortStateSuspend
  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
      If Vendor.DelFlag = 0 Then
        x.InsertRow = Vendor.VNAME & Chr$(9) & Vendor.vnum & Chr$(9) & VendorIdx.RecNum
      End If
  Next
  x.SortState = 1
  Close VendorIdxFile
  Close VendorFile
End Function

Public Function VendCodeList(x As fpCombo)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file

  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
      If Vendor.DelFlag = 0 Then
        If Vendor.ActiveFlag = 0 Then
          x.InsertRow = Vendor.vnum & Chr$(9) & VendorIdx.RecNum
        End If
      End If
  Next
  Close VendorIdxFile
  Close VendorFile
End Function
Public Function VendCodeNameIA(x As fpCombo)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file

  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
      If Vendor.DelFlag = 0 Then
        x.InsertRow = Vendor.vnum & Chr$(9) & Vendor.VNAME & Chr$(9) & VendorIdx.RecNum
      End If
  Next
  Close VendorIdxFile
  Close VendorFile
End Function

Public Function VendCodeName(x As fpCombo)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file

  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
      If Vendor.DelFlag = 0 Then
        If Vendor.ActiveFlag = 0 Then
          x.InsertRow = Vendor.vnum & Chr$(9) & Vendor.VNAME & Chr$(9) & VendorIdx.RecNum
        End If
      End If
  Next
  Close VendorIdxFile
  Close VendorFile
End Function
Public Function VendCodeName2(x As fpCombo, y As fpCombo)
  Dim VendorIdxFile As Integer, NumActiveVendors As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, cnt As Integer
  OpenVendorIdx VendorIdxFile, NumActiveVendors
  OpenVendorFile VendorFile, NumVRecs           'Open the Acct file

  For cnt = 1 To NumActiveVendors
    Get VendorIdxFile, cnt, VendorIdx
    Get VendorFile, VendorIdx.RecNum, Vendor
    x.Row = -1
    y.Row = -1
      If Vendor.DelFlag = 0 Then
        If Vendor.ActiveFlag = 0 Then
'use cnt instead of recnum for later range in invoice selection of payables
          x.InsertRow = Vendor.vnum & Chr$(9) & Vendor.VNAME & Chr$(9) & cnt
          y.InsertRow = Vendor.vnum & Chr$(9) & Vendor.VNAME & Chr$(9) & cnt
        End If
      End If
  Next
  Close VendorIdxFile
  Close VendorFile
End Function

Public Sub POList(x As fpCombo)
  Dim POEditFile As Integer, NumEdTrans As Integer, Transaction As Integer
  OpenPOEditFile POEditFile, NumEdTrans
  x.AddItem "All"
  For Transaction = 1 To NumEdTrans
    Get POEditFile, Transaction, POEdit
    If POEdit.Deleted <> True Then
      If Left$(POEdit.PONum, 3) <> "N/A" Then
        x.AddItem POEdit.PONum
      End If
    End If
  Next
  Close POEditFile
End Sub
Public Sub Post2PO(FileName$, BadTrans%, formname As Form, go4it As Boolean)
  Dim TrRecLen As Integer, File2Post As Integer, Num2Post As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, Log As String
  Dim TransFileNum As Integer, NumTrans As Long, cnt As Integer
  Dim POLogFileName As String, POLogFile As Integer, RecNum As Integer
  Dim Posted As Integer, Prev As Long, TransPosted As Integer
  Dim PRNFile As Integer, ReportFile As String
  Dim Acct As GLAcctRecType
  Dim Tran2Post As GLTransRecType        '--Dim a buffer for the edit file
 ' On Local Error GoTo ItsBroke

  TrRecLen = Len(Tran2Post)              'Determine the rec length
  File2Post = FreeFile                   'Get a handle on the Interface file
  Open FileName$ For Random As File2Post Len = TrRecLen
  Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions

  OpenAcctFile AcctFileNum, NumAccts     'Open & lock GL files
   'LOCK AcctFileNum

  OpenPOTransFile TransFileNum, NumTrans&
   'LOCK TransFileNum

   '--update the posting log file
  If go4it = True Then
    POLogFileName$ = "GLUTIL.LOG"
    POLogFile = FreeFile
    Open POLogFileName$ For Append As POLogFile
    Print #POLogFile, "Purchase Order initiated on " + Date$ + " @ " + Time$
    Log$ = Space$(132)
    'set correct Title for screen
    FrmShowPctComp.Label1 = "Posting PO Transactions"
  Else
    PRNFile = FreeFile
    ReportFile$ = "TempLog.PRN"
    Open ReportFile$ For Output As #PRNFile
    Print #PRNFile, "PO Verification initiated on " + Date$ + " @ " + Time$
    'set correct screen title
    FrmShowPctComp.Label1 = "Checking Accounts"
  End If
  FrmShowPctComp.CmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents
  For cnt = 1 To Num2Post                'Start processing transactions
    FrmShowPctComp.ShowPctComp cnt, Num2Post
    Get File2Post, cnt, Tran2Post

    RecNum = AcctFind(Tran2Post.AcctNum)   'Verify account is in G/L
    'Use recnum = 0 to test error log on verification
    'RecNum = 0
    If RecNum > 0 Then                  'if valid acct then proceed
    '''''If cnt = 25 Then Stop
         'tell user what's going on
        ' QPrintRC " Posting Account Number: ", 25, 1, 112
        ' QPrintRC Tran2Post.AcctNum, 25, 26, 112
'skip this part if posting potrans from invoices update acct.encumb there
'with correct amt from po not invoice.
         
         Get AcctFileNum, RecNum, Acct    'Get the account
         If Left$(Tran2Post.Src, 2) <> "AP" Then

         '--Update encumbrace field
         Select Case Acct.Typ
            Case "A", "E"                 'asset, exp accts
               Acct.Encumb = Round#(Acct.Encumb + Tran2Post.DrAmt - Tran2Post.CrAmt)
               If go4it = True Then
                 Put AcctFileNum, RecNum, Acct
               End If
            Case "L", "R"                 'liab, rev accts
               Acct.Encumb = Round#(Acct.Encumb + Tran2Post.CrAmt - Tran2Post.DrAmt)
               If go4it = True Then
                 Put AcctFileNum, RecNum, Acct
               End If
         End Select
         End If
         NumTrans& = NumTrans& + 1          'increment record pointer

         Get TransFileNum, NumTrans&, POTrans

         POTrans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
         POTrans.TRDATE = Tran2Post.TRDATE
         POTrans.Desc = Tran2Post.Desc
         POTrans.LDesc = Tran2Post.LDesc
         POTrans.CrAmt = Tran2Post.CrAmt
         POTrans.DrAmt = Tran2Post.DrAmt
         POTrans.Ref = Tran2Post.Ref
         POTrans.Src = Tran2Post.Src
         POTrans.NextTran = 0
         If go4it = True Then
           Put TransFileNum, NumTrans&, POTrans
         End If
         Posted = Posted + 1
         '---------------------------------Start linking here
         If Acct.FrstPTran = 0 Then        'if first trans for this acct,
            Acct.FrstPTran = NumTrans&      'assign first & last pointers to
            Acct.LastPTran = NumTrans&      'this transaction
            If go4it = True Then
              Put AcctFileNum, RecNum, Acct
            End If
         Else                             'otherwise
                                          'in the account file..
            Prev& = Acct.LastPTran             'remember the prev trans pointe
            Acct.LastPTran = NumTrans&        'reset last trans to this trans
            If go4it = True Then
              Put AcctFileNum, RecNum, Acct
            End If
                                          'In the POTrans file...
            Get TransFileNum, Prev&, POTrans    'Get the last transaction
            POTrans.NextTran = NumTrans&       'reset pointer to this trans
            If go4it = True Then
              Put TransFileNum, Prev&, POTrans
            End If
         End If

         TransPosted = TransPosted + 1

      Else                                'Account NOT found!
         BadTrans = BadTrans + 1          'Pass info back to caller
         '--how about an error log here.
         If go4it = True Then
            GoSub LogPOPostErr
         Else
            GoSub LogTempErr
         End If

      End If

   Next

   'UNLOCK AcctFileNum
   'UNLOCK TransFileNum
  If go4it = True Then
    If BadTrans = 0 Then
      Print #POLogFile, ("No Posting Errors. Posted Transaction Count: " + Using("####", TransPosted))
      Print #POLogFile, String$(80, "-")
    End If
  Else
    If BadTrans = 0 Then
      Print #PRNFile, ("No Errors Found. Transaction Count :" + Using$("####", TransPosted))
    End If
  End If
  Close AcctFileNum
  Close TransFileNum
  Close File2Post
  Close POLogFile
  Close
'Clean up editfile in calling program in case not posted
Exit Sub

POGotErr:
   Select Case Err
      Case 70
         Close
         MsgBox "Another user has the file locked, Please try again later.", vbOKOnly, "Access Denied"
         Exit Sub
      Case Else
   End Select
Return
LogTempErr:
   Print #PRNFile, "Error: Unpostable Transaction "
   Print #PRNFile, "Record Number  :"; Str$(cnt)
   Print #PRNFile, "Account Number :"; Tran2Post.AcctNum
   Print #PRNFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #PRNFile, "Description    :"; Tran2Post.Desc
   Print #PRNFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #PRNFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #PRNFile, "**********************"

Return

LogPOPostErr:
   Print #POLogFile, "Error: Unposted Transaction "
   Print #POLogFile, "Record Number  :"; Str$(cnt)
   Print #POLogFile, "Account Number :"; Tran2Post.AcctNum
   Print #POLogFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #POLogFile, "Description    :"; Tran2Post.Desc
   Print #POLogFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #POLogFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #POLogFile, "********************"

Return
ItsBroke:
  BadTrans = BadTrans + 1
  Print #PRNFile, "Error *** Call Software Support***"
  Print #PRNFile, "Record Number :"; Str$(cnt); Tran2Post.AcctNum
  Print #PRNFile, "Error Code"; Str(Err.Number)
  Resume Next

End Sub

Public Sub PrnOpenPays(formname As Form)
  Dim FF As String, MaxLines As Integer, Dash2 As String, Dash As String
  Dim PageNum As Integer, NumFunds As Integer, PrintFile As Integer
  Dim APDistRecLen As Integer, RecLen As Integer, VendorFile As Integer
  Dim NumVRecs As Integer, APLedgerFile As Integer, NumTran As Long
  Dim APDistFile As Integer, NumDistRecs As Long, VCnt As Integer
  Dim VendorIdxFile As Integer, NumVendors As Integer, Title As String
  Dim cnt As Integer, NumVendRecs As Long, DoneVHeader As Boolean
  Dim NextTrans As Long, Linecnt As Integer, fmt As String
  Dim ChkCnt As Integer, TotalChkAmt As Double, Page As String
  Dim VendTotal As Double, Vactive As Integer, fmt2 As String, fmTot As String
  Dim TPVend As String, ToPrint As String, TPInv As String
  Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
  FF$ = Chr$(12)
  MaxLines = 50
  Dash2$ = String$(80, "=")
  Dash$ = Dash2$
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  fmt$ = "###,###,###.##"
  fmt2$ = "#,###,###.##"
  fmTot$ = "####"
  TPVend$ = ""
  ToPrint$ = ""
  TPInv$ = ""
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub

  ReDim FundAmts(1 To NumFunds) As Double

  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  'ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  FrmShowPctComp.Label1 = "Processing Open Payables Report"
  FrmShowPctComp.Show , formname
  DoEvents
  EnableCloseButton formname.hwnd, False

  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  
  OpenVendorIdx VendorIdxFile, NumVendors
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  For cnt = 1 To NumVendRecs
  Get VendorIdxFile, cnt, VendorIdx
  '"apvendor.idx", VIndex(1), 12, NumVendRecs
    VIndex(cnt).RecNum = VendorIdx.RecNum
    VIndex(cnt).VendorCode = VendorIdx.VendorCode
  Next
  Close VendorIdxFile
  'DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))

  PrintFile = FreeFile

  Open "OPENPAYB.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen

 ' GoSub PrintOpenPayRptHeader

  For VCnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp VCnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DoneVHeader = 0
    Get VendorFile, VIndex(VCnt).RecNum, Vendor

    'IF VENDOR.DelFlag <> 0 THEN STOP
    NextTrans& = Vendor.FrstTran
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
        If Not DoneVHeader Then
'          If Linecnt > MaxLines Then
'            Print #PrintFile, FF$
'            Linecnt = 0
'          End If
          GoSub oPrintVendHeader
          'GoSub oPAInvHeader
        End If
        GoSub oPrintDist
'        If Linecnt > MaxLines Then
'          Print #PrintFile, FF$
'          GoSub PrintOpenPayRptHeader
'          GoSub oPrintVendHeader
'          GoSub oPAInvHeader
'        End If

      End If
      NextTrans& = APLedgerRec(1).NextTrans
    Loop

    If DoneVHeader Then
      GoSub oFinishVendor
    End If
'    If Linecnt > MaxLines Then
'      Print #PrintFile, FF$
'      GoSub PrintOpenPayRptHeader
'    End If
  Next

'  Print #PrintFile, 'JB
'  Print #PrintFile, Dash$
'  If Linecnt > MaxLines Then
'    Print #PrintFile, FF$
'  End If
'  GoSub FinishOpenReport
'  Print #PrintFile, FF$

  Close

  Erase FundList$, FundAmts, APLedgerRec, APDistRec

  Erase LedInfo, ChkRegInfo           ', ChkInfo

  Title$ = "Open Payables Report"
'  ViewPrint "OPENPAYB.PRN", title$
  EnableCloseButton formname.hwnd, True
  Load frmLoadingRpt
    ARptOpenPayable.GetName "OPENPAYB.PRN"
    ARptOpenPayable.txtTown.Caption = GLUserName$
    ARptOpenPayable.txtDate.Caption = Now
    ARptOpenPayable.Label1.Caption = "OPEN PAYABLES REPORT"
    ARptOpenPayable.totvends.DataValue = ChkCnt
    ARptOpenPayable.startrpt
oExitPreAudit:


  Exit Sub

'FinishOpenReport:
'  'PageNum = PageNum + 1
'  'Page$ = FUsing(STR$(PageNum), "###")
'  Print #PrintFile, "Report Totals:"
'  Print #PrintFile, "Vendors with Open Invoices: "; Using(fmTot, Str$(ChkCnt))
'  Print #PrintFile, "                  Totaling: "; Using(fmt$, Str$(TotalChkAmt#))
'  'PRINT #PrintFile,
'  'LineCnt = 7 who cares now?
'  Return

'PrintOpenPayRptHeader:
'  PageNum = PageNum + 1
'  Page$ = Using("###", Str$(PageNum))
'  Print #PrintFile, "A/P Open Payables Report                                       Page: " + Page$
'  Print #PrintFile, "Run Date: " + Date$
'  Linecnt = 2
'Return

'oPAInvHeader:
'  Print #PrintFile,             'Dash2$
'  Print #PrintFile, "Inv Date    Due Date    Inv Num                    PO                   Amount"
'  Print #PrintFile, "----------  ----------  -------------------------  ------------    ------------"
'  Linecnt = Linecnt + 3
'Return

oPrintDist:
  LSet ChkRegInfo(1).VendName = Vendor.VNAME
  LSet LedInfo(1).InvDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).DueDate = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).InvNum = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).PONum, 10)
  Else
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).MPONum, 10)
  End If
  RSet LedInfo(1).Amt = Using(fmt2$, Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  TPInv$ = LedInfo(1).InvNum + "~" + LedInfo(1).InvDate + "~" + LedInfo(1).DueDate
  TPInv$ = TPInv$ + "~" + LedInfo(1).PONum + "~" + LedInfo(1).Amt
  ToPrint$ = TPVend$ + "~" + TPInv$ + "~~~~~"
  Print #PrintFile, ToPrint$
'  Linecnt = Linecnt + 1

  'NextDist& = APLedgerRec(1).FrstDist
  'DO UNTIL NextDist& = 0
  '  GET APDistFile, NextDist&, APDistRec(1)
  '  IF LineCnt > MaxLines THEN
  '    PRINT #PrintFile, FF$
  '    GOSUB PrintOpenPayRptHeader
  '    'GOSUB PADistHeader
  '  END IF
  '  LSET LedInfo(1).InvDate = ""
  '  LSET LedInfo(1).DueDate = ""
  '  LSET LedInfo(1).InvNum = ""
  '  LSET LedInfo(1).PONum = ""
  '  RSET LedInfo(1).Amt = ""
  '  LSET LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
  '  RSET LedInfo(1).DistAmt = FUsing(STR$(APDistRec(1).DistAmt), ",########.#
  'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).Dist
  '  PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvN

  '  LineCnt = LineCnt + 1
  '  ThisFund$ = LEFT$(APDistRec(1).DistAcctNum, FundLen)
  '  FOR FundCnt = 1 TO NumFunds
  '    IF ThisFund$ = FundList$(FundCnt) THEN
  '      FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
  '      EXIT FOR
  '    END IF
  '  NEXT
  '  NextDist& = APDistRec(1).NextDist
  'LOOP
Return

oPrintVendHeader:
 ' Print #PrintFile,
 ' Print #PrintFile, Dash$
'  Print #PrintFile, Vendor.vnum; Vendor.VNAME
  DoneVHeader = -1
'  Linecnt = Linecnt + 3
  TPVend$ = QPTrim(Vendor.vnum) + "~" + QPTrim(Vendor.VNAME)
  Return

oFinishVendor:
'  Print #PrintFile, Tab(68); "------------"
'  Print #PrintFile, Tab(50); "Vendor Total: "; Tab(66); Using(fmt$, Str$(VendTotal#))
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
'  Linecnt = Linecnt + 2
  Vactive = 0
  Return
CancelExit:
  Exit Sub

End Sub

Public Sub PrnOpenPays2(formname As Form)
  Dim FF As String, MaxLines As Integer, Dash2 As String, Dash As String
  Dim PageNum As Integer, NumFunds As Integer, PrintFile As Integer
  Dim APDistRecLen As Integer, RecLen As Integer, VendorFile As Integer
  Dim NumVRecs As Integer, APLedgerFile As Integer, NumTran As Long
  Dim APDistFile As Integer, NumDistRecs As Long, VCnt As Integer
  Dim VendorIdxFile As Integer, NumVendors As Integer, Title As String
  Dim cnt As Integer, NumVendRecs As Long, DoneVHeader As Boolean
  Dim NextTrans As Long, Linecnt As Integer, fmt As String
  Dim ChkCnt As Integer, TotalChkAmt As Double, Page As String
  Dim VendTotal As Double, Vactive As Integer, fmt2 As String, fmTot As String
  FF$ = Chr$(12)
  MaxLines = 50
  Dash2$ = String$(80, "=")
  Dash$ = Dash2$
  Mid$(Dash2$, 1, 7) = Space$(7)
  PageNum = 0
  fmt$ = "###,###,###.##"
  fmt2$ = "#,###,###.##"
  fmTot$ = "####"
  ReDim FundList(1) As String
  GetFundList FundList$(), NumFunds
  If NumFunds = 0 Then Exit Sub

  ReDim FundAmts(1 To NumFunds) As Double

  ReDim APLedgerRec(1) As APLedger81RecType
  ReDim APDistRec(1) As APDistRecType
  ReDim LedInfo(1) As LedgerInfoType2
  'ReDim DistInfo(1) As DistInfoType
  ReDim ChkRegInfo(1) As CheckRegType
  FrmShowPctComp.Label1 = "Processing Open Payables Report"
  FrmShowPctComp.Show , formname
  DoEvents
  EnableCloseButton formname.hwnd, False

  NumVendRecs = (FileSize("apvendor.idx") \ 12)
  
  OpenVendorIdx VendorIdxFile, NumVendors
  ReDim VIndex(1 To NumVendRecs) As VendorIdxRecType
  For cnt = 1 To NumVendRecs
  Get VendorIdxFile, cnt, VendorIdx
  '"apvendor.idx", VIndex(1), 12, NumVendRecs
    VIndex(cnt).RecNum = VendorIdx.RecNum
    VIndex(cnt).VendorCode = VendorIdx.VendorCode
  Next
  Close VendorIdxFile
  'DistInfo(1).Fill1 = ""
  APDistRecLen = Len(APDistRec(1))
  RecLen = Len(APLedgerRec(1))

  PrintFile = FreeFile

  Open "OPENPAYB.PRN" For Output As PrintFile
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, RecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDistRecLen

  GoSub PrintOpenPayRptHeader

  For VCnt = 1 To NumVendRecs
    FrmShowPctComp.ShowPctComp VCnt, NumVendRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    DoneVHeader = 0
    Get VendorFile, VIndex(VCnt).RecNum, Vendor

    'IF VENDOR.DelFlag <> 0 THEN STOP
    NextTrans& = Vendor.FrstTran
    Do Until NextTrans& = 0
      Get APLedgerFile, NextTrans&, APLedgerRec(1)
      If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
        If Not DoneVHeader Then
          If Linecnt > MaxLines Then
            Print #PrintFile, FF$
            Linecnt = 0
          End If
          GoSub oPrintVendHeader
          GoSub oPAInvHeader
        End If
        GoSub oPrintDist
        If Linecnt > MaxLines Then
          Print #PrintFile, FF$
          GoSub PrintOpenPayRptHeader
          GoSub oPrintVendHeader
          GoSub oPAInvHeader
        End If

      End If
      NextTrans& = APLedgerRec(1).NextTrans
    Loop

    If DoneVHeader Then
      GoSub oFinishVendor
    End If
    If Linecnt > MaxLines Then
      Print #PrintFile, FF$
      GoSub PrintOpenPayRptHeader
    End If
  Next

  Print #PrintFile, 'JB
  Print #PrintFile, Dash$
  If Linecnt > MaxLines Then
    Print #PrintFile, FF$
  End If
  GoSub FinishOpenReport
  Print #PrintFile, FF$

  Close

  Erase FundList$, FundAmts, APLedgerRec, APDistRec

  Erase LedInfo, ChkRegInfo           ', ChkInfo

  Title$ = "Open Payables Report"
  ViewPrint "OPENPAYB.PRN", Title$
  EnableCloseButton formname.hwnd, True

oExitPreAudit:


  Exit Sub

FinishOpenReport:
  'PageNum = PageNum + 1
  'Page$ = FUsing(STR$(PageNum), "###")
  Print #PrintFile, "Report Totals:"
  Print #PrintFile, "Vendors with Open Invoices: "; Using(fmTot, Str$(ChkCnt))
  Print #PrintFile, "                  Totaling: "; Using(fmt$, Str$(TotalChkAmt#))
  'PRINT #PrintFile,
  'LineCnt = 7 who cares now?
  Return

PrintOpenPayRptHeader:
  PageNum = PageNum + 1
  Page$ = Using("###", Str$(PageNum))
  Print #PrintFile, "A/P Open Payables Report                                       Page: " + Page$
  Print #PrintFile, "Run Date: " + Date$
  Linecnt = 2
Return

oPAInvHeader:
  Print #PrintFile,             'Dash2$
  Print #PrintFile, "Inv Date   Due Date    Inv Num/Desc                      PO             Amount"
  Print #PrintFile, "---------- ----------  -------------------------     ------------  ------------"
  Linecnt = Linecnt + 3
Return

oPrintDist:
  LSet ChkRegInfo(1).VendName = Vendor.VNAME
  LSet LedInfo(1).InvDate = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).DueDate = Format(DateAdd("d", (APLedgerRec(1).DueDate), "12-31-1979"), "mm/dd/yyyy")
  LSet LedInfo(1).InvNum = QPTrim$(APLedgerRec(1).DOCNum) + "/" + QPTrim$(APLedgerRec(1).Comment)
  If Len(QPTrim$(APLedgerRec(1).PONum)) > 0 Then
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).PONum, 10)
  Else
    LSet LedInfo(1).PONum = Left$(APLedgerRec(1).MPONum, 10)
  End If
  RSet LedInfo(1).Amt = Using(fmt2$, Str$(APLedgerRec(1).Amt))
  VendTotal# = Round(VendTotal# + APLedgerRec(1).Amt)
  TotalChkAmt# = Round(TotalChkAmt# + APLedgerRec(1).Amt)
  Print #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvNum; LedInfo(1).PONum; Tab(68); LedInfo(1).Amt
  Linecnt = Linecnt + 1

  'NextDist& = APLedgerRec(1).FrstDist
  'DO UNTIL NextDist& = 0
  '  GET APDistFile, NextDist&, APDistRec(1)
  '  IF LineCnt > MaxLines THEN
  '    PRINT #PrintFile, FF$
  '    GOSUB PrintOpenPayRptHeader
  '    'GOSUB PADistHeader
  '  END IF
  '  LSET LedInfo(1).InvDate = ""
  '  LSET LedInfo(1).DueDate = ""
  '  LSET LedInfo(1).InvNum = ""
  '  LSET LedInfo(1).PONum = ""
  '  RSET LedInfo(1).Amt = ""
  '  LSET LedInfo(1).DistAcct = APDistRec(1).DistAcctNum
  '  RSET LedInfo(1).DistAmt = FUsing(STR$(APDistRec(1).DistAmt), ",########.#
  'PRINT #PrintFile, DistInfo(1).Fill1; DistInfo(1).DistAcct; DistInfo(1).Dist
  '  PRINT #PrintFile, LedInfo(1).InvDate; LedInfo(1).DueDate; LedInfo(1).InvN

  '  LineCnt = LineCnt + 1
  '  ThisFund$ = LEFT$(APDistRec(1).DistAcctNum, FundLen)
  '  FOR FundCnt = 1 TO NumFunds
  '    IF ThisFund$ = FundList$(FundCnt) THEN
  '      FundAmts(FundCnt) = Round(FundAmts(FundCnt) + APDistRec(1).DistAmt)
  '      EXIT FOR
  '    END IF
  '  NEXT
  '  NextDist& = APDistRec(1).NextDist
  'LOOP
Return

oPrintVendHeader:
  Print #PrintFile,
  Print #PrintFile, Dash$
  Print #PrintFile, Vendor.vnum; Vendor.VNAME
  DoneVHeader = -1
  Linecnt = Linecnt + 3
  Return

oFinishVendor:
  Print #PrintFile, Tab(68); "------------"
  Print #PrintFile, Tab(50); "Vendor Total: "; Tab(66); Using(fmt$, Str$(VendTotal#))
  VendTotal# = 0
  ChkCnt = ChkCnt + 1
  Linecnt = Linecnt + 2
  Vactive = 0
  Return
CancelExit:
  Exit Sub

End Sub
  
    
Public Sub updateaccttots()
Dim GLAcct    As GLAcctRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FYStartDate As Integer
Dim ActiveYear As Integer
  Dim EndDate As Integer
  Dim Acct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim YTDBal As Double
  Dim Account As String
  Dim YTDSum As Double

On Local Error GoTo Goterr
   GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
   EndDate = DateDiff("d", "12/31/1979", Now)
    If EndDate >= FY2BegDate Then
      ActiveYear = 2
      FYStartDate = FY2BegDate
    Else
      ActiveYear = 1
      FYStartDate = FY1BegDate
    End If

  FixPOEncumbRpt EndDate, FYStartDate

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  For Acct = 1 To NumGLAccts
    YTDBal# = 0
    Get AcctIdxFileNum, Acct, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct
      '--We want only revenue or expenditure accounts
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then
        NextTr& = GLAcct.FrstTran 'get the first trans for this acct
        Do Until NextTr& = 0    'keep going 'til we run out
          Get TransFileNum, NextTr&, GLTrans
          If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then
            Select Case GLAcct.Typ
            Case "E"
              YTDBal# = Round#(YTDBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            Case "R"
              YTDBal# = Round#(YTDBal# + GLTrans.CrAmt - GLTrans.DrAmt)
            End Select
          End If
          NextTr& = GLTrans.NextTran              'Get the next transaction
        Loop
        '--Put the new totals in the file
        GLAcct.YTD = Round#(YTDBal#)
        Put AcctFileNum, GLAcctidx.RecNum, GLAcct
      End If    '--test for rev or exp accts
  Next          'Process next account
Goterr:

  Close

End Sub
