Attribute VB_Name = "BLConvertCommon"
Option Explicit
  Public ScreenW As Long
  Public coladj As Double
  Public InvalidBal() As Double
  Public InvalidName() As String
  Public InvalidNum() As String
  Public InvalidCnt As Integer
  Public doAlign As Boolean
  Public alnRpt$
  Public StartPath As String
  Public BadMaskFlag As Boolean
  Public NumOfAligns As Integer
  Public DupNums() As Integer
  Public DupBlanks() As Integer
  Public NextDup As Integer
  Public NextBlank As Integer
  Public NonNums() As Integer
  Public NonCnt As Integer
  Public DifNums() As Integer
  Public NumOfDifs As Integer
  Public UMillion() As Integer
  Public NumOfUS As Integer
  Public OMillion() As Integer
  Public NumOfOS As Integer
  Public DupCustCats() As Integer
  Public DCCCnt As Integer
  Public DifBalCnt As Integer
  Public DifBalRecs() As Integer
  Public DifBalTAmt() As Double
  Public DifBalCAmt() As Double
  Public Version As Integer
  Public CatVersion As Integer
  Public DupCats() As Integer
  Public DupCatCnt As Integer
  Public BlankCatDesc() As Integer
  Public BlankCatDescCnt As Integer
  
  Public Const AcctFileName = "GLACCT.DAT"
  Public Const BLCatCodeName = "ARCODE.DAT"
  Public Const TransFileName = "GLTRANS.DAT"
  Public Const CatCodeIdxName = "arcatcodeidx.dat"
  Public Const LicNumIdx = "arlicnumidx.dat"
  Public Const CustNameIdx = "arcustnameidx.dat"
  Public Const CustNumIdx = "arcustnumidx.dat"
  Public Const CustSearchNameIdx = "arsrhidx.dat"
  Public Const BLCustFileName = "ARCUST.DAT"
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
Public Sub OpenCatCodeIdxFile(CatCodeIdxHandle As Integer)
  Dim CatCodeIdx As CatCodeIdxType
  Dim CatCodeIdxLen As Integer
  
  CatCodeIdxLen = Len(CatCodeIdx)
  CatCodeIdxHandle = FreeFile
  Open CatCodeIdxName For Random Shared As CatCodeIdxHandle Len = CatCodeIdxLen
End Sub
Public Sub OpenCatCodeFile(CatCodeHandle As Integer)
  Dim CatCodeRec As ARNewCatCodeRecType
  Dim CatCodeLen As Integer
  CatCodeLen = Len(CatCodeRec)
  CatCodeHandle = FreeFile
  Open BLCatCodeName For Random Shared As CatCodeHandle Len = CatCodeLen
End Sub
Public Sub OpenSrchNameIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustSearchNameIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustSearchNameIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenLicNumIdxFile(LicIdxHandle As Integer)
  Dim LicIdx As CustLicNumIdxType
  Dim LicIdxLen As Integer
  LicIdxLen = Len(LicIdx)
  LicIdxHandle = FreeFile
  Open LicNumIdx For Random Shared As LicIdxHandle Len = LicIdxLen
End Sub

Public Sub OpenCustNumIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustNumIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustNumIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenDosCatFile2(CatHandle As Integer)
  Dim CatRec As DosARNewCatCodeRecType2
  Dim CatLen As Integer
  CatLen = Len(CatRec)
  CatHandle = FreeFile
  Open BLCatCodeName For Random Shared As CatHandle Len = CatLen
End Sub
Public Sub OpenDosCatFile(CatHandle As Integer)
  Dim CatRec As DosARNewCatCodeRecType
  Dim CatLen As Integer
  CatLen = Len(CatRec)
  CatHandle = FreeFile
  Open BLCatCodeName For Random Shared As CatHandle Len = CatLen
End Sub
Public Sub OpenCustNameIdxFile(CustIdxHandle As Integer)
  Dim CustIdx As CustNameIdxType
  Dim CustIdxLen As Integer
  CustIdxLen = Len(CustIdx)
  CustIdxHandle = FreeFile
  Open CustNameIdx For Random Shared As CustIdxHandle Len = CustIdxLen
End Sub
Public Sub OpenDosCustFile2(CustHandle As Integer)
  Dim CustRec As DOSARCustRecType2
  Dim CustLen As Integer
  CustLen = Len(CustRec)
  CustHandle = FreeFile
  Open BLCustFileName For Random Shared As CustHandle Len = CustLen
End Sub
Public Sub OpenCustFile(CustHandle As Integer)
  Dim CustRec As ARCustRecType
  Dim CustLen As Integer
  CustLen = Len(CustRec)
  CustHandle = FreeFile
  Open BLCustFileName For Random Shared As CustHandle Len = CustLen
End Sub
Public Sub OpenDosCustFile(CustHandle As Integer)
  Dim CustRec As DosARCustRecType
  Dim CustLen As Integer
  CustLen = Len(CustRec)
  CustHandle = FreeFile
  Open BLCustFileName For Random Shared As CustHandle Len = CustLen
End Sub
Public Sub OpenGLAcctFile(GLHandle As Integer)
  Dim GLRec As GLAcctRecType
  Dim GLRecLen As Integer
  GLRecLen = Len(GLRec)
  GLHandle = FreeFile
  Open AcctFileName For Random Shared As GLHandle Len = GLRecLen
End Sub

Public Function QPTrim$(Text As String)
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
Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then 'added 7/24
    Kill FileName$
  End If
End Sub
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

Public Function Exist(FileName$) As Boolean
  Dim FileHandle As Integer
  Dim TempSize As Long
  On Local Error Resume Next
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  TempSize = LOF(FileHandle)
  Close FileHandle
  If TempSize <= 0 Then
    Kill FileName$
    Exist = False
  Else
    Exist = True
  End If

End Function

Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
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
                  If Year > 1919 And Year < 2099 Then
                      CheckValDate = True
                  End If
              End If
          End If
      End If
End Function

Public Function Date2Num%(TheDate$)
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function
Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function

Public Sub CreateLicNumIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigLic As Double
  Dim ThisLic As Double
  Dim ThisX As Integer
  Dim SmallLic As Double
  Dim TempLic As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim LicIdx As CustLicNumIdxType
  Dim LicIdxHandle As Integer
  Dim LicIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustLicNumIdxType
  Dim ThisCnt As Integer
  
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempLicIdx(1 To NumOfCustRecs) As CustLicNumIdxType
  
  BigLic = 999999999999#
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
'    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo BadNum
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempLicIdx(ThisCnt).CustRec = x
    TempLicIdx(ThisCnt).LicNum = QPTrim$(CustRec.LICENSE)
    ThisLic = Val(CustRec.LICENSE)
    If ThisLic > BigLic Then
      BigLic = ThisLic
    End If
BadNum:
  Next x
  Close CustHandle
  
'  BigLic = BigLic + "A"
  SmallLic = BigLic
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt ' NumOfCustRecs
      ThisLic = Val(TempLicIdx(x).LicNum)
      If ThisLic < SmallLic Then
        SmallLic = ThisLic
        ThisX = x
      End If
    Next x
    HoldThis = TempLicIdx(Nextx)
    TempLicIdx(Nextx) = TempLicIdx(ThisX)
    TempLicIdx(ThisX) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallLic = BigLic
  Loop
  
  OpenLicNumIdxFile LicIdxHandle
  For x = 1 To ThisCnt ' NumOfCustRecs
    LicIdx = TempLicIdx(x)
    Put LicIdxHandle, x, LicIdx
  Next x
  Close LicIdxHandle
  
  
End Sub

Public Sub CreateCatCodeIdx()
  Dim BigNum$ ' As Double
  Dim ThisNum$ ' As Double
  Dim ThisX As Integer
  Dim SmallNum$ ' As Double
  Dim x As Integer
  Dim Nextx As Integer
  Dim TempNum As Integer
  Dim CodeHandle As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeRecLen As Integer
  Dim NumOfCodeRecs As Integer
  Dim CodeIdx As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CatCodeIdxType

  OpenCatCodeFile CodeHandle

  NumOfCodeRecs = LOF(CodeHandle) \ Len(CodeRec)
  ReDim TempCodeIdx(1 To NumOfCodeRecs) As CatCodeIdxType

  BigNum = 0
  For x = 1 To NumOfCodeRecs
    Get CodeHandle, x, CodeRec

    TempCodeIdx(x).CatCodeRec = x
    TempCodeIdx(x).CatCodeNum = QPTrim$(CodeRec.CATCODE)
    ThisNum = QPTrim$(CodeRec.CATCODE)
    If Val(ThisNum) > Val(BigNum) Then
      BigNum = ThisNum
    End If
BadNum:
  Next x
  Close CodeHandle

  BigNum = BigNum + "1"
  SmallNum = BigNum
  Nextx = 1

  Do
    For x = Nextx To NumOfCodeRecs
      ThisNum = TempCodeIdx(x).CatCodeNum
      If Val(ThisNum) < Val(SmallNum) Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempCodeIdx(Nextx)
    TempCodeIdx(Nextx) = TempCodeIdx(ThisX)
    TempCodeIdx(ThisX) = HoldThis
    If Nextx = NumOfCodeRecs Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop

  OpenCatCodeIdxFile CodeIdxHandle
  For x = 1 To NumOfCodeRecs
    CodeIdx = TempCodeIdx(x)
    Put CodeIdxHandle, x, CodeIdx
  Next x

  Close

End Sub
Public Sub CreateCustNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigName$
  Dim ThisName$
  Dim ThisX As Integer
  Dim SmallName$
  Dim TempName As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustNameIdxType
  Dim ThisCnt As Integer
  
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
'    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo BadNum
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).BillingName = QPTrim$(CustRec.BILLNAME)
    ThisName = QPTrim$(CustRec.BILLNAME)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt ' NumOfCustRecs
      ThisName = TempCustIdx(x).BillingName
      If ThisName < SmallName Then
        SmallName = ThisName
        ThisX = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(ThisX)
    TempCustIdx(ThisX) = HoldThis
    If Nextx = ThisCnt Then Exit Do ' NumOfCustRecs Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
  Loop
  
  OpenCustNameIdxFile CustIdxHandle
  For x = 1 To ThisCnt ' NumOfCustRecs
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  
  
End Sub

Public Sub CreateCustNumIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigNum As Integer
  Dim ThisNum As Integer
  Dim ThisX As Integer
  Dim SmallNum As Integer
  Dim TempNum As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustNumIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustNumIdxType
  Dim ThisCnt As Integer
  
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustNumIdxType
  
  BigNum = 0
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
'    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo ItsDeleted
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).CUSTNUMB = CInt(CustRec.CUSTNUMB)
    ThisNum = CInt(CustRec.CUSTNUMB)
    If ThisNum > BigNum Then
      BigNum = ThisNum
    End If
ItsDeleted:
  Next x
  Close CustHandle
  
  SmallNum = BigNum + 1
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisNum = CInt(TempCustIdx(x).CUSTNUMB)
      If ThisNum < SmallNum Then
        SmallNum = ThisNum
        ThisX = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(ThisX)
    TempCustIdx(ThisX) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  
  OpenCustNumIdxFile CustIdxHandle
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  
End Sub

Public Sub CreateCustSearchNameIdx()
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim n As Integer
  Dim Nextx As Integer
  Dim y As Integer, cnt As Integer
  Dim ThisText$, CustRecNo As Integer
  Dim CustCnt As Integer
  Dim BigName$
  Dim ThisName$
  Dim ThisX As Integer
  Dim SmallName$
  Dim TempName As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustIdx As CustSearchNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecLen As Integer
  Dim RecNum As Integer
  Dim HoldThis As CustSearchNameIdxType
  Dim ThisCnt As Integer
  
  OpenCustFile CustHandle
  
  NumOfCustRecs = LOF(CustHandle) \ Len(CustRec)
  ReDim TempCustIdx(1 To NumOfCustRecs) As CustSearchNameIdxType
  
  BigName = "A"
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo BadNum
    ThisCnt = ThisCnt + 1
    TempCustIdx(ThisCnt).CustRec = x
    TempCustIdx(ThisCnt).SORTNAME = QPTrim$(CustRec.SORTNAME)
    ThisName = QPTrim$(CustRec.SORTNAME)
    If ThisName > BigName Then
      BigName = ThisName
    End If
BadNum:
  Next x
  Close CustHandle
  
  BigName = BigName + "A"
  SmallName = BigName
  Nextx = 1
  
  Do
    For x = Nextx To ThisCnt
      ThisName = TempCustIdx(x).SORTNAME
      If ThisName < SmallName Then
        SmallName = ThisName
        ThisX = x
      End If
    Next x
    HoldThis = TempCustIdx(Nextx)
    TempCustIdx(Nextx) = TempCustIdx(ThisX)
    TempCustIdx(ThisX) = HoldThis
    If Nextx = ThisCnt Then Exit Do
    Nextx = Nextx + 1
    SmallName = BigName
  Loop
  
  OpenSrchNameIdxFile CustIdxHandle
  For x = 1 To ThisCnt
    CustIdx = TempCustIdx(x)
    Put CustIdxHandle, x, CustIdx
  Next x
  Close CustIdxHandle
  
End Sub

'Public Sub ReLinkTransactions()
'  Dim TransRec As ARTransRecType
'  Dim TranFile As Integer
'  Dim NumCRec&
'  Dim CustFile As Integer
'  Dim ARCust As ARCustRecType
'  Dim NumTRec&
'  Dim Ccnt&
'  Dim TCnt&
'  Dim CustRec&
'  Dim BadTran As Integer
'
'  ReDim ARTran(1 To 2) As ARTransRecType
'
'  OpenCustFile CustFile
'
'  NumCRec& = LOF(CustFile) / Len(ARCust)
'
'  OpenTransFile TranFile
'  NumTRec& = LOF(TranFile) / Len(TransRec)
'
'  For Ccnt& = 1 To NumCRec&
'    Get CustFile, Ccnt&, ARCust
'    ARCust.FirstTrans = 0
'    ARCust.LastTrans = 0
'    Put CustFile, Ccnt&, ARCust
'  Next
'
'  For TCnt& = 1 To NumTRec&
'    Get TranFile, TCnt&, ARTran(1)
'    CustRec = Val(ARTran(1).CustomerNumber)
'    If (CustRec& > 0) And (CustRec& <= NumCRec&) Then
'      Get CustFile, CustRec, ARCust
'      If ARCust.LastTrans = 0 Then
'        ARCust.FirstTrans = TCnt&
'        ARCust.LastTrans = TCnt&
'        Put CustFile, CustRec&, ARCust
'        ARTran(1).NextTrans = 0
'        Put TranFile, TCnt&, ARTran(1)
'      Else
'        Get TranFile, ARCust.LastTrans, ARTran(2)  'get old last tr
'        ARTran(2).NextTrans = TCnt&                    'point it to next tr
'        Put TranFile, ARCust.LastTrans, ARTran(2)  'put it back
'        ARCust.LastTrans = TCnt&                    'set new cust last TR
'        Put CustFile, CustRec&, ARCust          'put it back
'        ARTran(1).NextTrans = 0
'        Put TranFile, TCnt&, ARTran(1)
'      End If
'    Else
'    End If
'NoGood:
'  Next
'
'End Sub

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
  
End Function

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmBLViewPrintConv.ReportName = ReportFile$
   frmBLViewPrintConv.Caption = Title
   frmBLViewPrintConv.PgNum = PgNum
   frmBLViewPrintConv.cmdAlignment.Visible = False
   If ForceSBar Then
     frmBLViewPrintConv.fpMemo1.ScrollBars = BothFixed
   Else
     frmBLViewPrintConv.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmBLViewPrintConv.cmdAlignment.Enabled = True
     frmBLViewPrintConv.AlignRpt = AlgnRptfile$
    Else
      frmBLViewPrintConv.cmdAlignment.Enabled = False
    End If
   frmBLViewPrintConv.Show 1
   doAlign = False
End Sub
Public Function GetGLRecNum(GLNum$) As Long
  Dim NumOfGLRecs As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim x As Long
  
  GetGLRecNum = 0
  OpenGLAcctFile AcctHandle
  NumOfGLRecs = LOF(AcctHandle) / Len(GLAcctRec)
  For x = 1 To NumOfGLRecs
    Get AcctHandle, x, GLAcctRec
    If QPTrim$(GLAcctRec.Num) = GLNum Then
      GetGLRecNum = x
      Exit For
    End If
  Next x
  Close AcctHandle

End Function

Public Function Check4DupLicNums() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim y As Integer
  Dim RunCnt As Integer
  Dim ChangeFlag As Boolean
  Dim HoldThis$
  Dim TempHold As Integer
  Dim ThisCnt As Integer
  Dim ThisX As Integer
  Dim ThisRecCnt As Integer
  Dim BigLic As Double
  Dim DupCnt As Integer
  Dim ThisPct As Double
  Dim NextLic As Integer
  Dim NextRec As Integer
  Dim GotIt As Boolean
  
  If NextDup > 0 Then
    ReDim DupNums(1 To 1)
    NextDup = 0
  End If
  
  ChangeFlag = False
  RunCnt = 1
  GotIt = False
  ReDim HoldDups(1 To 1) As String
  Check4DupLicNums = False 'none found
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
    DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
    DosNumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
  End If
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  ElseIf DosNumOfCustRecs = 1 Then
    Close
    Exit Function
  End If
  ReDim ThisIdx(1 To 1) As Integer
  
  For x = 1 To DosNumOfCustRecs
    If Version = 1 Then
      Get DosCustHandle, x, DosCustRec
      If QPTrim$(DosCustRec.LICENSE) = "" Then GoTo ItsDeleted
      If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      Nextx = Nextx + 1
      ReDim Preserve ThisIdx(1 To Nextx) As Integer
      ThisIdx(Nextx) = x
      ThisCnt = ThisCnt + 1
    ElseIf Version = 2 Then
'      If x = 52 Then Stop
      Get DosCustHandle2, x, DosCustRec2
      If QPTrim$(DosCustRec2.LICENSE) = "" Then GoTo ItsDeleted
      If QPTrim$(DosCustRec2.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      Nextx = Nextx + 1
      ReDim Preserve ThisIdx(1 To Nextx) As Integer
      ThisIdx(Nextx) = x
      ThisCnt = ThisCnt + 1
    End If
ItsDeleted:
  Next x

  Nextx = 1
  ReDim LicFound(1 To 1) As String
  
  If Version = 1 Then
    Do
      Get DosCustHandle, Nextx, DosCustRec
      HoldThis = QPTrim$(DosCustRec.LICENSE)
      For x = Nextx + 1 To ThisCnt
        Get DosCustHandle, ThisIdx(x), DosCustRec
        If NextLic > 0 Then
          If GotIt = False Then
            For y = 1 To NextLic 'check against license already known to be dups
            'then if found there is no need to keep looking
              If HoldThis = LicFound(y) Then GoTo NotThisOne
            Next y
          End If
        End If
        If HoldThis = QPTrim(DosCustRec.LICENSE) Then
          If GotIt = False Then
            NextLic = NextLic + 1
            ReDim Preserve LicFound(1 To NextLic) As String
            LicFound(NextLic) = QPTrim(DosCustRec.LICENSE)
            NextDup = NextDup + 2
            ReDim Preserve DupNums(1 To NextDup) As Integer
            DupNums(NextDup - 1) = Nextx
            DupNums(NextDup) = ThisIdx(x)
            GotIt = True
          Else
            NextDup = NextDup + 1
            ReDim Preserve DupNums(1 To NextDup) As Integer
            DupNums(NextDup) = ThisIdx(x)
          End If
        End If
NotThisOne:
      Next x
      GotIt = False
      Nextx = Nextx + 1
      If Nextx > ThisCnt Then Exit Do
      RunCnt = RunCnt + 1
      ThisPct = OldRound(RunCnt / ThisCnt) * 100
      frmBLConvertMain.fptxtMarque = "Looking for duplicate license numbers is " + CStr(ThisPct) + "% completed"
    Loop
  End If
  
  If Version = 2 Then
    Do
      Get DosCustHandle2, Nextx, DosCustRec2
      HoldThis = QPTrim$(DosCustRec2.LICENSE)
      For x = Nextx + 1 To ThisCnt
        Get DosCustHandle2, ThisIdx(x), DosCustRec2
        If NextLic > 0 Then
          If GotIt = False Then
            For y = 1 To NextLic 'check against license already known to be dups
            'then if found there is no need to keep looking
              If HoldThis = LicFound(y) Then GoTo NotThisOne2
            Next y
          End If
        End If
        If HoldThis = QPTrim(DosCustRec2.LICENSE) Then
          If GotIt = False Then
            NextLic = NextLic + 1
            ReDim Preserve LicFound(1 To NextLic) As String
            LicFound(NextLic) = QPTrim(DosCustRec2.LICENSE)
            NextDup = NextDup + 2
            ReDim Preserve DupNums(1 To NextDup) As Integer
            DupNums(NextDup - 1) = Nextx
            DupNums(NextDup) = ThisIdx(x)
            GotIt = True
          Else
            NextDup = NextDup + 1
            ReDim Preserve DupNums(1 To NextDup) As Integer
            DupNums(NextDup) = ThisIdx(x)
          End If
        End If
NotThisOne2:
      Next x
      GotIt = False
      Nextx = Nextx + 1
      If Nextx > ThisCnt Then Exit Do
      RunCnt = RunCnt + 1
      ThisPct = OldRound(RunCnt / ThisCnt) * 100
      frmBLConvertMain.fptxtMarque = "Looking for duplicate license numbers is " + CStr(ThisPct) + "% completed"
    Loop
  End If
  
  If Version = 1 Then
    Close DosCustHandle
  ElseIf Version = 2 Then
    Close DosCustHandle2
  End If
  
  If NextDup > 0 Then
    Check4DupLicNums = True
  End If

End Function
Public Function Check4CustDupCats() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim x As Integer, y As Integer
  Dim Nextx As Integer
  Dim MultiCats As Integer
  
  Check4CustDupCats = False
  OpenDosCustFile DosCustHandle
  DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  End If
  
  ReDim CustCatCodes(1 To 5) As String
  ReDim DupCustCats(1 To 1) As Integer
  DCCCnt = 0
  For x = 1 To DosNumOfCustRecs
      Get DosCustHandle, x, DosCustRec
      If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      MultiCats = 0
      CustCatCodes(1) = QPTrim$(DosCustRec.BILLCAT1)
      If CustCatCodes(1) <> "" Then MultiCats = MultiCats + 1
      CustCatCodes(2) = QPTrim$(DosCustRec.BILLCAT2)
      If CustCatCodes(2) <> "" Then MultiCats = MultiCats + 1
      CustCatCodes(3) = QPTrim$(DosCustRec.BILLCAT3)
      If CustCatCodes(3) <> "" Then MultiCats = MultiCats + 1
      CustCatCodes(4) = QPTrim$(DosCustRec.BILLCAT4)
      If CustCatCodes(4) <> "" Then MultiCats = MultiCats + 1
      CustCatCodes(5) = QPTrim$(DosCustRec.BILLCAT5)
      If CustCatCodes(5) <> "" Then MultiCats = MultiCats + 1
      If MultiCats <= 1 Then GoTo ItsDeleted
      Nextx = 0
      Do
        Nextx = Nextx + 1
        If Nextx > MultiCats Then Exit Do
        For y = 1 To MultiCats
          If y = Nextx Then GoTo NotThisOne
          If CustCatCodes(Nextx) = CustCatCodes(y) And CustCatCodes(y) <> "" Then
            DCCCnt = DCCCnt + 1
            ReDim Preserve DupCustCats(1 To DCCCnt) As Integer
            DupCustCats(DCCCnt) = x
            Check4CustDupCats = True
            GoTo ItsDeleted
          End If
NotThisOne:
        Next y
      Loop
ItsDeleted:
  Next x

  Close DosCustHandle

End Function
Public Function Check4NonNums() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer
  Dim Nextx As Integer
  
  If NonCnt > 0 Then
    ReDim NonNums(1 To 1)
    NonCnt = 0
  End If
  
  NonCnt = 0
  Check4NonNums = False 'none found
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
    DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
    DosNumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
  End If
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  ElseIf DosNumOfCustRecs = 1 Then
    Close
    Exit Function
  End If
  
  For x = 1 To DosNumOfCustRecs
    If Version = 1 Then
      Get DosCustHandle, x, DosCustRec
      If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      If Not IsNumeric(DosCustRec.CUSTNUMB) Then
        Nextx = Nextx + 1
        ReDim Preserve NonNums(1 To Nextx) As Integer
        NonNums(Nextx) = x
        NonCnt = NonCnt + 1
      End If
    ElseIf Version = 2 Then
      Get DosCustHandle2, x, DosCustRec2
      If QPTrim$(DosCustRec2.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      If Not IsNumeric(DosCustRec2.CUSTNUMB) Then
        Nextx = Nextx + 1
        ReDim Preserve NonNums(1 To Nextx) As Integer
        NonNums(Nextx) = x
        NonCnt = NonCnt + 1
      End If
    End If
ItsDeleted:
  Next x

  If Version = 1 Then
    Close DosCustHandle
  ElseIf Version = 2 Then
    Close DosCustHandle2
  End If
  
  If NonCnt > 0 Then
    Check4NonNums = True
  End If

End Function
Public Function Check4DifBalAmts() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim x As Integer, Nextx As Integer
  Dim RunCnt As Integer
  Dim ThisCnt As Integer
  
  If NumOfDifs > 0 Then
    ReDim DifNums(1 To 1)
    NumOfDifs = 0
  End If
  
  Check4DifBalAmts = False 'none found
  
  OpenDosCustFile DosCustHandle
  DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  End If
  
  ReDim ThisIdx(1 To 1) As Integer
  For x = 1 To DosNumOfCustRecs 'weed out the deleted
    Get DosCustHandle, x, DosCustRec
    If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
    If DosCustRec.AcctBal < -1000000 Or DosCustRec.AcctBal > 1000000 Then GoTo ItsDeleted
    Nextx = Nextx + 1
    ReDim Preserve ThisIdx(1 To Nextx) As Integer
    ThisIdx(Nextx) = x
    ThisCnt = ThisCnt + 1
ItsDeleted:
  Next x
  
  For x = 1 To ThisCnt
    Get DosCustHandle, ThisIdx(x), DosCustRec
      If OldRound(DosCustRec.PenBal + DosCustRec.LicBal) <> OldRound(DosCustRec.AcctBal) Then
        RunCnt = RunCnt + 1
        ReDim Preserve DifNums(1 To RunCnt) As Integer
        DifNums(RunCnt) = ThisIdx(x)
      End If
  Next x
  
  
  NumOfDifs = RunCnt
  If RunCnt > 0 Then
    Check4DifBalAmts = True
  End If
End Function

Public Function Check4OvrUndMillion() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer, Nextx As Integer
  Dim RunCnt As Integer
  Dim ThisCnt As Integer
  Dim ThisBal$
  
  If NumOfUS > 0 Then
    ReDim UMillion(1 To 1)
    NumOfUS = 0
  End If
  
  If NumOfOS > 0 Then
    ReDim OMillion(1 To 1)
    NumOfOS = 0
  End If
  
  Check4OvrUndMillion = False 'none found
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
    DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
    DosNumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
  End If
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  End If
  
  For x = 1 To DosNumOfCustRecs 'weed out the deleted
    If Version = 1 Then
      Get DosCustHandle, x, DosCustRec
      If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      ThisBal = CStr(DosCustRec.AcctBal)
      If InStr(ThisBal, "E") Then 'GoTo ItsDeleted
        If InStr(ThisBal, "-") Then
          NumOfUS = NumOfUS + 1
          ReDim Preserve UMillion(1 To NumOfUS) As Integer
          UMillion(NumOfUS) = x
        Else
          NumOfOS = NumOfOS + 1
          ReDim Preserve OMillion(1 To NumOfOS) As Integer
          OMillion(NumOfOS) = x
        End If
      End If
    ElseIf Version = 2 Then
      Get DosCustHandle2, x, DosCustRec2
      If QPTrim$(DosCustRec2.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      ThisBal = CStr(DosCustRec2.AcctBal)
      If InStr(ThisBal, "E") Then 'GoTo ItsDeleted
        If InStr(ThisBal, "-") Then
          NumOfUS = NumOfUS + 1
          ReDim Preserve UMillion(1 To NumOfUS) As Integer
          UMillion(NumOfUS) = x
        Else
          NumOfOS = NumOfOS + 1
          ReDim Preserve OMillion(1 To NumOfOS) As Integer
          OMillion(NumOfOS) = x
        End If
      End If
    End If
ItsDeleted:
  Next x
  
  Close
  If NumOfOS > 0 Or NumOfUS > 0 Then
    Check4OvrUndMillion = True
  End If
  
End Function

Public Function Check4BlankLicNums() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosNumOfCustRecs2 As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer
  
  If NextBlank > 0 Then
    ReDim DupBlanks(1 To 1)
    NextBlank = 0
  End If
  
  Check4BlankLicNums = False 'none found
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
    DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
    DosNumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
  End If
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  End If
  
  For x = 1 To DosNumOfCustRecs
    If Version = 1 Then
      Get DosCustHandle, x, DosCustRec
      If QPTrim$(DosCustRec.LICENSE) <> "" Then GoTo ItsDeleted
      If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      NextBlank = NextBlank + 1
      ReDim Preserve DupBlanks(1 To NextBlank) As Integer
      DupBlanks(NextBlank) = x
    ElseIf Version = 2 Then
      Get DosCustHandle2, x, DosCustRec2
      If QPTrim$(DosCustRec2.LICENSE) <> "" Then GoTo ItsDeleted
      If QPTrim$(DosCustRec2.SORTNAME) = "DELETED" Then GoTo ItsDeleted
      NextBlank = NextBlank + 1
      ReDim Preserve DupBlanks(1 To NextBlank) As Integer
      DupBlanks(NextBlank) = x
    End If
ItsDeleted:
  Next x
  
  If Version = 1 Then
    Close DosCustHandle
  ElseIf Version = 2 Then
    Close DosCustHandle2
  End If
  
  If NextBlank > 0 Then
    Check4BlankLicNums = True
  End If
  
End Function
  
Public Sub ChangeCat2Nums()
  Dim NumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer, y As Integer
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim NumOfCodeRecs As Integer
  Dim CodeNum As Integer
  Dim ThisCode As Integer
  Dim Changed1 As Boolean
  Dim Changed2 As Boolean
  Dim Changed3 As Boolean
  Dim Changed4 As Boolean
  Dim Changed5 As Boolean
  Dim ThisPct As Integer
  
  'this code examines category codes...if codes are letters and not
  'numbers then this code changes them to numbers
  
  frmBLMessageBoxJrWOpts.Label1.Caption = "The category version to be changed is Version #" + CStr(CatVersion) + ". Click 'F10' to continue with the category change procedure."
  frmBLMessageBoxJrWOpts.Label1.Top = 700
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
    DoEvents
  End If
  
  frmBLMessageBoxJrWOpts.Label1.Caption = "The customer version to be updated is Version #" + CStr(Version) + ". Click 'F10' to continue with the customer update procedure."
  frmBLMessageBoxJrWOpts.Label1.Top = 700
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
    DoEvents
  End If
  
  If CatVersion = 1 Then
    CodeNum = 10000
  Else
    CodeNum = 100
  End If
  
  If Version = 1 Then
    OpenDosCustFile DosCustHandle
    NumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
    If NumOfCustRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no customers saved. Procedure aborted."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      Close DosCustHandle
      Exit Sub
    End If
  ElseIf Version = 2 Then
    OpenDosCustFile2 DosCustHandle2
    NumOfCustRecs = LOF(DosCustHandle2) / Len(DosCustRec2)
    If NumOfCustRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no customers saved. Procedure aborted."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      Close DosCustHandle2
      Exit Sub
    End If
  End If
  
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
    NumOfCodeRecs = LOF(DosCodeHandle) \ Len(DosCodeRec)
    If NumOfCodeRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no codes saved. Procedure aborted."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      Close DosCodeHandle
      Exit Sub
    End If
  ElseIf CatVersion = 2 Then
    OpenDosCatFile2 DosCodeHandle2
    NumOfCodeRecs = LOF(DosCodeHandle2) \ Len(DosCodeRec2)
    If NumOfCodeRecs = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no codes saved. Procedure aborted."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      Close DosCodeHandle2
      Exit Sub
    End If
  End If
  
  ReDim CodeBefore(1 To NumOfCodeRecs) As String * 5
  ReDim CodeAfter(1 To NumOfCodeRecs) As String * 5
  If CatVersion = 1 Then
    For x = 1 To NumOfCodeRecs
      Get DosCodeHandle, x, DosCodeRec
        CodeBefore(x) = QPTrim$(DosCodeRec.CATCODE)
        CodeAfter(x) = CStr(CodeNum)
        DosCodeRec.CATCODE = CStr(CodeNum)
        Put DosCodeHandle, x, DosCodeRec
        CodeNum = CodeNum + 1
    Next x
  ElseIf CatVersion = 2 Then
    For x = 1 To NumOfCodeRecs
      Get DosCodeHandle2, x, DosCodeRec2
        CodeBefore(x) = QPTrim$(DosCodeRec2.CATCODE)
        CodeAfter(x) = CStr(CodeNum)
        DosCodeRec2.CATCODE = CStr(CodeNum)
        Put DosCodeHandle2, x, DosCodeRec2
        CodeNum = CodeNum + 1
    Next x
  End If
  
  If Version = 1 And CatVersion = 1 Then
    For x = 1 To NumOfCustRecs
      Changed1 = False
      Changed2 = False
      Changed3 = False
      Changed4 = False
      Changed5 = False
      Get DosCustHandle, x, DosCustRec
      For y = 1 To NumOfCodeRecs
      Get DosCodeHandle, y, DosCodeRec
        If Changed1 = False Then
          If QPTrim$(DosCustRec.BILLCAT1) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT1 = QPTrim$(DosCodeRec.CATCODE)
            Changed1 = True
          End If
        End If
        If Changed2 = False Then
          If QPTrim$(DosCustRec.BILLCAT2) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT2 = QPTrim$(DosCodeRec.CATCODE)
            Changed2 = True
          End If
        End If
        If Changed3 = False Then
          If QPTrim$(DosCustRec.BILLCAT3) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT3 = QPTrim$(DosCodeRec.CATCODE)
            Changed3 = True
          End If
        End If
        If Changed4 = False Then
          If QPTrim$(DosCustRec.BILLCAT4) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT4 = QPTrim$(DosCodeRec.CATCODE)
            Changed4 = True
          End If
        End If
        If Changed5 = False Then
          If QPTrim$(DosCustRec.BILLCAT5) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT5 = QPTrim$(DosCodeRec.CATCODE)
            Changed5 = True
          End If
        End If
      Next y
      If Changed1 = False Then
        DosCustRec.BILLCAT1 = ""
      End If
      If Changed2 = False Then
        DosCustRec.BILLCAT2 = ""
      End If
      If Changed3 = False Then
        DosCustRec.BILLCAT3 = ""
      End If
      If Changed4 = False Then
        DosCustRec.BILLCAT4 = ""
      End If
      If Changed5 = False Then
        DosCustRec.BILLCAT5 = ""
      End If
      Put DosCustHandle, x, DosCustRec
      ThisPct = OldRound((x / NumOfCustRecs) * 100)
      frmBLConvertMain.fptxtMarque.Text = "Changing category codes is " + CStr(ThisPct) + "% completed."
    Next x
    Close 'no other files opened now
  End If
  
  If Version = 2 And CatVersion = 2 Then
    For x = 1 To NumOfCustRecs
      Changed1 = False
      Changed2 = False
      Changed3 = False
      Changed4 = False
      Changed5 = False
      Get DosCustHandle2, x, DosCustRec2
      For y = 1 To NumOfCodeRecs
      Get DosCodeHandle2, y, DosCodeRec2
        If Changed1 = False Then
          If QPTrim$(DosCustRec2.BILLCAT) = QPTrim$(CodeBefore(y)) Then
            DosCustRec2.BILLCAT = QPTrim$(DosCodeRec2.CATCODE)
            Changed1 = True
          End If
        End If
      Next y
      Put DosCustHandle2, x, DosCustRec2
      ThisPct = OldRound((x / NumOfCustRecs) * 100)
      frmBLConvertMain.fptxtMarque.Text = "Changing category codes is " + CStr(ThisPct) + "% completed."
    Next x
    Close
  End If

  If Version = 1 And CatVersion = 2 Then
    For x = 1 To NumOfCustRecs
      Changed1 = False
      Changed2 = False
      Changed3 = False
      Changed4 = False
      Changed5 = False
      Get DosCustHandle, x, DosCustRec
      For y = 1 To NumOfCodeRecs
      Get DosCodeHandle2, y, DosCodeRec2
        If Changed1 = False Then
          If QPTrim$(DosCustRec.BILLCAT1) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT1 = QPTrim$(DosCodeRec2.CATCODE)
            Changed1 = True
          End If
        End If
        If Changed2 = False Then
          If QPTrim$(DosCustRec.BILLCAT2) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT2 = QPTrim$(DosCodeRec2.CATCODE)
            Changed2 = True
          End If
        End If
        If Changed3 = False Then
          If QPTrim$(DosCustRec.BILLCAT3) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT3 = QPTrim$(DosCodeRec2.CATCODE)
            Changed3 = True
          End If
        End If
        If Changed4 = False Then
          If QPTrim$(DosCustRec.BILLCAT4) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT4 = QPTrim$(DosCodeRec2.CATCODE)
            Changed4 = True
          End If
        End If
        If Changed5 = False Then
          If QPTrim$(DosCustRec.BILLCAT5) = QPTrim$(CodeBefore(y)) Then
            DosCustRec.BILLCAT5 = QPTrim$(DosCodeRec2.CATCODE)
            Changed5 = True
          End If
        End If
      Next y
      If Changed1 = False Then
        DosCustRec.BILLCAT1 = ""
      End If
      If Changed2 = False Then
        DosCustRec.BILLCAT2 = ""
      End If
      If Changed3 = False Then
        DosCustRec.BILLCAT3 = ""
      End If
      If Changed4 = False Then
        DosCustRec.BILLCAT4 = ""
      End If
      If Changed5 = False Then
        DosCustRec.BILLCAT5 = ""
      End If
      Put DosCustHandle, x, DosCustRec
      ThisPct = OldRound((x / NumOfCustRecs) * 100)
      frmBLConvertMain.fptxtMarque.Text = "Changing category codes is " + CStr(ThisPct) + "% completed."
    Next x
    Close 'no other files opened now
  End If

  If Version = 2 And CatVersion = 1 Then
    For x = 1 To NumOfCustRecs
      Changed1 = False
      Changed2 = False
      Changed3 = False
      Changed4 = False
      Changed5 = False
      Get DosCustHandle2, x, DosCustRec2
      For y = 1 To NumOfCodeRecs
      Get DosCodeHandle, y, DosCodeRec
        If Changed1 = False Then
          If QPTrim$(DosCustRec2.BILLCAT) = QPTrim$(CodeBefore(y)) Then
            DosCustRec2.BILLCAT = QPTrim$(DosCodeRec.CATCODE)
            Changed1 = True
          End If
        End If
      Next y
      Put DosCustHandle2, x, DosCustRec2
      ThisPct = OldRound((x / NumOfCustRecs) * 100)
      frmBLConvertMain.fptxtMarque.Text = "Changing category codes is " + CStr(ThisPct) + "% completed."
    Next x
    Close
  End If
  
  Close
  
  frmBLMessageBoxJr.Label1.Caption = "SUCCESS: Changing category codes to numbers and updating customer category data has completed. Please run version check again to make sure the new data is accurate."
  frmBLMessageBoxJr.Label1.Top = 700
  frmBLMessageBoxJr.Show vbModal
  DoEvents
  frmBLConvertMain.fptxtMarque.Text = "Begin Conversion"
  frmBLConvertMain.cmdChangeCat2Nums.Enabled = False
End Sub

Public Sub Check4DupCats()
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer, y As Integer, z As Integer
  Dim ThisCode$
  Dim FoundDup As Boolean
  
  'this sub looks for duplicate category codes...
  'if a category has a duplicate then when alpha codes
  'are changed to numeric codes then any customer who is
  'using the duplicate code will have their code changed
  'to match the first duplicate even if the duplicate
  'category code does not match the description (it probably
  'wouldn't) of the category they were using...duplicate codes
  'must be fixed before continuing
  DupCatCnt = 0
  ReDim DupCats(1 To 1) As Integer
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
    NumOfCodeRecs = LOF(DosCodeHandle) \ Len(DosCodeRec)
    If NumOfCodeRecs = 0 Then
      Close DosCodeHandle
      Exit Sub
    End If
    For x = 1 To NumOfCodeRecs
      FoundDup = False
      Get DosCodeHandle, x, DosCodeRec
      ThisCode = QPTrim$(DosCodeRec.CATCODE)
      For y = 1 To NumOfCodeRecs
        Get DosCodeHandle, y, DosCodeRec
        If y = x Then GoTo CheckedAlready
          If DupCatCnt > 0 Then
            For z = 1 To DupCatCnt
              If DupCats(z) < 0 Then GoTo CheckedAlready
              If y = DupCats(z) Then GoTo CheckedAlready
            Next z
          End If
          If QPTrim$(DosCodeRec.CATCODE) = ThisCode Then
            If FoundDup = False Then
              DupCatCnt = DupCatCnt + 1
              ReDim Preserve DupCats(1 To DupCatCnt) As Integer
              DupCats(DupCatCnt) = -x
              FoundDup = True
            End If
            DupCatCnt = DupCatCnt + 1
            ReDim Preserve DupCats(1 To DupCatCnt) As Integer
            DupCats(DupCatCnt) = y
          End If
CheckedAlready:
      Next y
    Next x
  ElseIf CatVersion = 2 Then
    OpenDosCatFile2 DosCodeHandle2
    NumOfCodeRecs = LOF(DosCodeHandle2) \ Len(DosCodeRec2)
    If NumOfCodeRecs = 0 Then
      Close DosCodeHandle2
      Exit Sub
    End If
    For x = 1 To NumOfCodeRecs
      FoundDup = False
      Get DosCodeHandle2, x, DosCodeRec2
      ThisCode = QPTrim$(DosCodeRec2.CATCODE)
      For y = 1 To NumOfCodeRecs
        Get DosCodeHandle2, y, DosCodeRec2
        If y = x Then GoTo CheckedAlready2
          If DupCatCnt > 0 Then
            For z = 1 To DupCatCnt
              If DupCats(z) < 0 Then GoTo CheckedAlready2
              If y = DupCats(z) Then GoTo CheckedAlready2
            Next z
          End If
          If QPTrim$(DosCodeRec2.CATCODE) = ThisCode Then
            If FoundDup = False Then
              DupCatCnt = DupCatCnt + 1
              ReDim Preserve DupCats(1 To DupCatCnt) As Integer
              DupCats(DupCatCnt) = -x
              FoundDup = True
            End If
            DupCatCnt = DupCatCnt + 1
            ReDim Preserve DupCats(1 To DupCatCnt) As Integer
            DupCats(DupCatCnt) = y
          End If
CheckedAlready2:
      Next y
    Next x
  End If
  Close
End Sub

Public Function Check4BlankCatDesc() As Boolean
  Dim DosNumOfCustRecs As Integer
  Dim DosCustRec As DosARCustRecType
  Dim DosCustHandle As Integer
  Dim DosCustRec2 As DOSARCustRecType2
  Dim DosCustHandle2 As Integer
  Dim x As Integer
  
  Check4BlankCatDesc = False
  BlankCatDescCnt = 0
  
  ReDim BlankCatDesc(1 To 1)
  BlankCatDescCnt = 0
  
  OpenDosCustFile DosCustHandle
  DosNumOfCustRecs = LOF(DosCustHandle) / Len(DosCustRec)
  
  If DosNumOfCustRecs = 0 Then
    MsgBox "No customers on file."
    Close
    Exit Function
  End If
  
  For x = 1 To DosNumOfCustRecs
    Get DosCustHandle, x, DosCustRec
    If QPTrim$(DosCustRec.SORTNAME) = "DELETED" Or QPTrim$(DosCustRec.Deleted) = "Y" Then GoTo ItsDeleted
    If QPTrim$(DosCustRec.BILLCAT1) <> "" Then
      If QPTrim$(DosCustRec.DESC1) = "" Then
        BlankCatDescCnt = BlankCatDescCnt + 1
        ReDim Preserve BlankCatDesc(1 To BlankCatDescCnt) As Integer
        BlankCatDesc(BlankCatDescCnt) = x
        GoTo ItsDeleted
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT2) <> "" Then
      If QPTrim$(DosCustRec.DESC2) = "" Then
        BlankCatDescCnt = BlankCatDescCnt + 1
        ReDim Preserve BlankCatDesc(1 To BlankCatDescCnt) As Integer
        BlankCatDesc(BlankCatDescCnt) = x
        GoTo ItsDeleted
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT3) <> "" Then
      If QPTrim$(DosCustRec.DESC3) = "" Then
        BlankCatDescCnt = BlankCatDescCnt + 1
        ReDim Preserve BlankCatDesc(1 To BlankCatDescCnt) As Integer
        BlankCatDesc(BlankCatDescCnt) = x
        GoTo ItsDeleted
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT4) <> "" Then
      If QPTrim$(DosCustRec.DESC4) = "" Then
        BlankCatDescCnt = BlankCatDescCnt + 1
        ReDim Preserve BlankCatDesc(1 To BlankCatDescCnt) As Integer
        BlankCatDesc(BlankCatDescCnt) = x
        GoTo ItsDeleted
      End If
    End If
    If QPTrim$(DosCustRec.BILLCAT5) <> "" Then
      If QPTrim$(DosCustRec.DESC5) = "" Then
        BlankCatDescCnt = BlankCatDescCnt + 1
        ReDim Preserve BlankCatDesc(1 To BlankCatDescCnt) As Integer
        BlankCatDesc(BlankCatDescCnt) = x
        GoTo ItsDeleted
      End If
    End If
ItsDeleted:
  Next x

  Close DosCustHandle
  
  If BlankCatDescCnt > 0 Then
    Check4BlankCatDesc = True
  End If

End Function

Public Function Clear4CatCodeProblems() As Integer
  Dim DosCodeRec As DosARNewCatCodeRecType
  Dim DosCodeHandle As Integer
  Dim DosCodeRec2 As DosARNewCatCodeRecType2
  Dim DosCodeHandle2 As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer, y As Integer, z As Integer
  Dim NoDescCnt As Integer
  Dim NoCodeCnt As Integer
  
  Clear4CatCodeProblems = 0
  ReDim NoDesc(1 To 1) As Integer
  ReDim NoCode(1 To 1) As Integer
  NoDescCnt = 0
  NoCodeCnt = 0
  If CatVersion = 1 Then
    OpenDosCatFile DosCodeHandle
    NumOfCodeRecs = LOF(DosCodeHandle) \ Len(DosCodeRec)
    If NumOfCodeRecs = 0 Then
      Close DosCodeHandle
      Exit Function
    End If
    
    For x = 1 To NumOfCodeRecs
      Get DosCodeHandle, x, DosCodeRec
        If QPTrim$(DosCodeRec.CODEDESC) = "" Then
          NoDescCnt = NoDescCnt + 1
          ReDim Preserve NoDesc(1 To NoDescCnt) As Integer
          NoDesc(NoDescCnt) = x
        End If
        If QPTrim$(DosCodeRec.CATCODE) = "" Then
          NoCodeCnt = NoCodeCnt + 1
          ReDim Preserve NoCode(1 To NoCodeCnt) As Integer
          NoCode(NoCodeCnt) = x
        End If
    Next x
    Close DosCodeHandle
  End If
    
  If NoDescCnt > 0 And NoCodeCnt > 0 Then
    Clear4CatCodeProblems = 3
  ElseIf NoDescCnt > 0 Then
    Clear4CatCodeProblems = 1
  ElseIf NoCodeCnt > 0 Then
    Clear4CatCodeProblems = 2
  End If
      
End Function

Public Function GetCatDesc(CatNum$) As String
  Dim x As Integer
  Dim CatRec As ARNewCatCodeRecType
  Dim CHandle  As Integer
  Dim CatRecNums As Integer
  
  GetCatDesc = ""
  OpenCatCodeFile CHandle
  CatRecNums = LOF(CHandle) / Len(CatRec)
    
  If CatRecNums = 0 Then Exit Function
  
  For x = 1 To CatRecNums
    Get CHandle, x, CatRec
      If QPTrim$(CatRec.CATCODE) = QPTrim$(CatNum$) Then
        GetCatDesc = QPTrim$(CatRec.CODEDESC)
        Exit For
      End If
  Next x
  
  Close CHandle
  
End Function
