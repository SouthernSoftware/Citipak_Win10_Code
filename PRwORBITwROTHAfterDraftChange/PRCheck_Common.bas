Attribute VB_Name = "PRCheck_Common"
Option Explicit
  Public RecNum As Long
  Public Emp2Rec(1) As EmpData2Type
  Public EHandle As Integer
  Public TRHandle As Integer
  Public BadMaskFlag As Boolean
  Public EntryType As Integer
  Public ScreenW As Long
  Public coladj As Double
  Public doAlign As Boolean
  Public alnRpt As String
  Public CancelDoAlign As Boolean
  Public ChkPrintOn As Boolean
  Public ReprintChkOn As Boolean
  Public ComputerName As String
  Public ToPrint1(1 To 10) As Integer
  Public ToPrint2(1 To 10) As Integer
  Public InFileNames(1 To 10) As String
  Public OutFileNames(1 To 10) As String
  Public StartPath As String
  Public NumOfAligns As Integer '7/23
  Public RegExit As Boolean
  Public RptOpt As Integer
  Public Twiddle As String
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
  
            Public Const PRData = "prdata\"
   Public Const ErnCodeFileName = "PRERNCOD.DAT"
      Public Const UnitFileName = "PRUNIT.DAT"
       Public Const SysFileName = "PRSYS.DAT"
 Public Const TransWorkFileName = "PRTRANST.DAT"
    Public Const CheckPrintFile = "CHKPRNT.DAT"
      Public Const EmpData2Name = "PREMP2.DAT"
      Public Const EmpData3Name = "PREMP3.DAT"
       Public Const EmpIdxLName = "PREMPL.IDX"  'name idx
       Public Const EmpIdxNName = "PREMPN.IDX"  'numb idx
    Public Const AccrueFileName = "PRACCRUE.DAT"
     Public Const LeaveFileName = "PRLEAVE.DAT"
    Public Const ChecksFileName = "PRCHECKS.DAT"
 Public Const PPDefaultFileName = "PRPPDEF.DAT"
   Public Const DedCodeFileName = "PRDEDCOD.DAT"
   Public Const PRDraftFileName = "PRDRAFTI.DAT"
   Public Const EmpDataFileMask = "PRRPTS\PREMPRPT.DPM"
    Public Const GLAcctIdxFile = "BAACCTDX.DAT"
    Public Const JGLAcctIdxFile = "GLACCT.IDX"
  Public Const PrinterSetUpFile = "PRPRNSET.DAT"
      Public Const PRActiveFile = "PRDATA\PRACTIVE.FLG"    '*
      Public Const TempVoidFileName = "PRDATA\TEMPVOID.DAT"
Public Const PPDraftInfoFileName = "PRDATA\PPDFINFO.DAT"
Public Sub Main()
  Dim ShellHandle As Integer
  Dim XHandle As Integer
  Dim Cnt2$

  If Exist("prdata\frompr.dat") Then 'pulls the user's password
  'number saved from payroll.exe and saves it in the global variable
  'PWcnt
    ShellHandle = FreeFile
    Open "prdata\frompr.dat" For Input As ShellHandle
    Line Input #ShellHandle, Cnt2
    PWcnt = CInt(Cnt2)
    Close ShellHandle 'this .dat file is deleted now from payroll.exe
  End If
  
  KillFile ("prdata\frompr.dat")
  frmChkPrintingMenu.Show
  DoEvents
End Sub
Public Sub OpenTempVoidFile(TempVoidHandle As Integer) '12/12/02
  Dim TempVoidLen As Integer
  Dim TempVoid As VoidCheckType
  TempVoidLen = Len(TempVoid)
  TempVoidHandle = FreeFile
  Open TempVoidFileName For Random Shared As TempVoidHandle Len = TempVoidLen
End Sub
Public Sub OpenLeaveFileName(LeaveHandle As Integer)
  Dim LeaveRec As LeaveRecType
  Dim LeaveRecLen As Integer
  LeaveRecLen = Len(LeaveRec)
  LeaveHandle = FreeFile
  Open PRData + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
End Sub
Public Sub OpenPrinterSetupFile(PrinterSUFHandle As Integer)
  Dim PrinterSUFRec As PRNSetupRecType
  Dim PrinterSUFRecLen As Integer
  PrinterSUFRecLen = Len(PrinterSUFRec)
  PrinterSUFHandle = FreeFile
  Open PRData + PrinterSetUpFile For Random Shared As PrinterSUFHandle Len = PrinterSUFRecLen
End Sub
Public Sub OpenPPDraftInfo(PPDraftInfoHandle As Integer)
  Dim PPDraftInfoRec As GLAcctIndexType
  Dim PPDraftInfoRecLen As Integer
  PPDraftInfoRecLen = Len(PPDraftInfoRec)
  PPDraftInfoHandle = FreeFile
  Open PPDraftInfoFileName For Random Shared As PPDraftInfoHandle Len = PPDraftInfoRecLen
End Sub
Public Sub OpenChecksFile(ChecksHandle As Integer)
  Dim ChecksRec As PRCheckRecType
  Dim ChecksRecLen As Integer
  ChecksRecLen = Len(ChecksRec)
  ChecksHandle = FreeFile
  Open PRData + ChecksFileName For Random Shared As ChecksHandle Len = ChecksRecLen
End Sub

Public Sub OpenPRChecksFile(PRChecksHandle As Integer)
  Dim PRChecksRec As PRCheckRecType
  Dim PRChecksRecLen As Integer
  PRChecksRecLen = Len(PRChecksRec)
  PRChecksHandle = FreeFile
  Open PRData + ChecksFileName For Random Shared As PRChecksHandle Len = PRChecksRecLen
End Sub

Public Sub OpenTransWorkFile(TransWorkFileHandle As Integer)
  Dim TransWorkFileRec As TransRecType
  Dim TransWorkRecLen As Integer
  TransWorkRecLen = Len(TransWorkFileRec)
  TransWorkFileHandle = FreeFile
  Open PRData + TransWorkFileName For Random Shared As TransWorkFileHandle Len = TransWorkRecLen
End Sub
'****************************************************************************
'OldRounds a double precision value to nearest hundredth
'****************************************************************************
Public Function OldRound#(n As Double)
'  OldRound# = Round(n, 2)
  OldRound# = Int(n * 100 + 0.5) / 100
End Function
   
Public Sub OpenEmpIdxLNameFile(EmpIdxLNameHandle As Integer)
  EmpIdxLNameHandle = FreeFile
  Open PRData + EmpIdxLName For Random Shared As EmpIdxLNameHandle Len = 2
End Sub
   
Public Sub OpenEmpIdxNNameFile(EmpIdxNNameHandle As Integer)
  EmpIdxNNameHandle = FreeFile
  Open PRData + EmpIdxNName For Random Shared As EmpIdxNNameHandle Len = 2
End Sub
Public Sub OpenEmpData3File(EmpData3FileHandle As Integer)
  Dim EmpData3FileRec As EmpData3Type
  Dim EmpData3RecLen As Integer
  EmpData3RecLen = Len(EmpData3FileRec)
  EmpData3FileHandle = FreeFile
  Open PRData + EmpData3Name For Random Shared As EmpData3FileHandle Len = EmpData3RecLen
End Sub
Public Sub OpenEmpData2File(EmpData2FileHandle As Integer)
  Dim EmpData2FileRec As EmpData2Type
  Dim EmpData2RecLen As Integer
  EmpData2RecLen = Len(EmpData2FileRec)
  EmpData2FileHandle = FreeFile
  Open PRData + EmpData2Name For Random Shared As EmpData2FileHandle Len = EmpData2RecLen
End Sub
   
Public Sub OpenDedCodeFile(DedCodeFileHandle As Integer)
  Dim DedCodeFileRec As DedCodeRecType
  Dim DedCodeRecLen As Integer
  DedCodeRecLen = Len(DedCodeFileRec)
  DedCodeFileHandle = FreeFile
  Open PRData + DedCodeFileName For Random Shared As DedCodeFileHandle Len = DedCodeRecLen
End Sub
   
Public Sub OpenErnCodeFile(ErnCodeFileHandle As Integer)
  Dim ErnCodeFileRec As ErnCodeRecType
  Dim ErnCodeRecLen As Integer
  ErnCodeRecLen = Len(ErnCodeFileRec)
  ErnCodeFileHandle = FreeFile
  Open PRData + ErnCodeFileName For Random Shared As ErnCodeFileHandle Len = ErnCodeRecLen
End Sub
   
Public Sub OpenPRDraftFile(PRDraftFileHandle As Integer)
  Dim PRDraftFileRec As DraftInfoFileName
  Dim PRDraftRecLen As Integer
  PRDraftRecLen = Len(PRDraftFileRec)
  PRDraftFileHandle = FreeFile
  Open PRData + PRDraftFileName For Random Shared As PRDraftFileHandle Len = PRDraftRecLen
End Sub
   
Public Sub OpenSysFile(SysFileHandle As Integer)
  Dim SysFileRec As RegDSysFileRecType
  Dim SysRecLen As Integer
  SysRecLen = Len(SysFileRec)
  SysFileHandle = FreeFile
  Open PRData + SysFileName For Random Shared As SysFileHandle Len = SysRecLen
End Sub

Public Sub OpenUnitFile(FileHandle As Integer)
  Dim UnitFileRec As UnitFileRecType
  Dim UnitRecLen As Integer
  UnitRecLen = Len(UnitFileRec)
  FileHandle = FreeFile
  Open PRData + UnitFileName For Random Shared As FileHandle Len = UnitRecLen
End Sub

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
Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim NewText As String
  Dim thischar$
  Dim CTChar$
  Dim TTChar$
  Dim CTLen As Integer
  Dim TTLen As Integer
  Dim BigLen As Integer
  
  StrLen = Len(Text)
  CTLen = Len(ChangeThis$)
  TTLen = Len(ToThis$)
  If CTLen > TTLen Then
    BigLen = CTLen
  ElseIf TTLen > CTLen Then
    BigLen = TTLen
  Else
    BigLen = CTLen
  End If
  
  For cnt = 1 To StrLen
    thischar = Mid$(Text, cnt, 1)
    CTChar = Mid$(Text, cnt, CTLen)
    If CTChar = ChangeThis Then
      NewText = NewText + ToThis
      cnt = cnt + BigLen
    Else
      NewText = NewText + thischar
    End If
  Next
  ReplaceString$ = Trim$(NewText)
End Function
Public Sub KillFile(FileName As String)
  If Exist(FileName$) Then
    Kill FileName$
  End If
End Sub

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
   
'This function is a replacement for the QuickPak FileSize function.
'Due to the way Windows NT updates a file's size in the directory, an
'error can occur using DOS Function 4Eh (Find first file service) to
'read a file's size from the Directory. You can force Windows NT to
'commit the directory info by just opening the file again.
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

Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmLoadingRpt.Show
   frmViewPrintChks.ReportName = ReportFile$
   frmViewPrintChks.Caption = Title
   frmViewPrintChks.PgNum = PgNum
   frmViewPrintChks.cmdAlignment.Visible = False
   If ForceSBar Then
     frmViewPrintChks.fpMemo1.ScrollBars = BothFixed
   Else
     frmViewPrintChks.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmViewPrintChks.cmdAlignment.Enabled = True
     frmViewPrintChks.AlignRpt = AlgnRptfile$
    Else
      frmViewPrintChks.cmdAlignment.Enabled = False
    End If
   frmViewPrintChks.Show 1
   Unload frmLoadingRpt
   doAlign = False
End Sub
 

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

Public Function Date2Num%(TheDate$)
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function
Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

Function CheckFor2ManyDecimals(Text As String) As Boolean
  Dim cnt As Integer
  Dim DecCnt As Integer
  Dim StrLen As Long
  Dim thischar$
  
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    thischar = Mid$(Text, cnt, 1)
    If thischar = "." Then DecCnt = DecCnt + 1
  Next cnt
  If DecCnt > 1 Then
    CheckFor2ManyDecimals = True
  Else
    CheckFor2ManyDecimals = False
  End If
End Function

'this function returns a pointer into
'the employee index array. whose record we restart
'check printing with
Function GetStartEmp%(FirstBadChkNum&)
  Dim GRecNum&
  Dim CheckRecLen As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&, GotEmp As Boolean
  Dim NumOfRecs As Long
  Dim CHandle As Integer, cnt As Long
  Dim IdxNHandle As Integer
  
  If FirstBadChkNum& <= 0 Then GoTo SkipChkScrh

  ReDim Check(1) As PRCheckRecType
  CheckRecLen = Len(Check(1))
  IdxRecLen = 2
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) \ IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For cnt = 1 To NumOfRecs
    Get IdxNHandle, cnt, IdxBuff(cnt)
  Next cnt
  Close IdxNHandle
  OpenChecksFile CHandle
  For cnt = 1 To NumOfRecs
    GRecNum& = CLng(IdxBuff(cnt))
    Get CHandle, GRecNum&, Check(1)
    If Check(1).CheckNum = FirstBadChkNum& Then
      GotEmp = True
      Exit For
    End If
  Next

  Close CHandle

SkipChkScrh:
  If GotEmp Then
    GetStartEmp% = cnt
  Else
    GetStartEmp% = 0
  End If


End Function

Sub PrintChecks(FirstEmp, StartChkNum&, Num2Print, LPTPort, CheckNum, CheckDate&, ReprintFlag As Boolean)

  Dim FF$, Title$, TCheckNum&
  Dim CHandle As Integer
  Dim PPDFHandle As Integer
  Dim DedHandle As Integer
  Dim ErnHandle As Integer
  Dim THandle As Integer
  Dim EHandle2 As Integer
  Dim EHandle3 As Integer
  Dim RHandle As Integer
  Dim TransRecLen As Long
  Dim Emp2RecLen As Long
  Dim Emp3RecLen As Long
  Dim CRecLen As Long
  Dim CheckRecLen As Long
  Dim PPDFLen As Long
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Long
  Dim PPDFFile As Integer
  Dim HOLBAL#, Cnt2 As Integer
  Dim PayType$, AllHrsPaid#
  Dim DoneCnt As Long, GRecNum&, VacPay#
  Dim YGross#, HRate#, WorkPay#, SickPay#, CompPay#, HolPay#, PerPay#
  Dim x As Integer, PERBAL#
  Dim DedCnt As Integer
  Dim DedRec As DedCodeRecType
  Dim IdxNHandle As Integer
  Dim CheckPrintFileName$, PPDraftInfoFileName$, cnt As Integer
  Dim FreqCnt As Integer
  Dim PayFreq, Pg(1) As String
  Dim IdxNRec As NumbSortIdxType
  ReDim FreqPay$(1 To 7)
  Dim Page As Integer, UTemp$
  Dim Unit(1) As UnitFileRecType
  Dim UHandle As Integer
  Dim LineCnt As Integer
  Dim CPFNHandle As Integer
  Dim CheckStyle As Integer
  Dim Nextx As Integer
  Dim DedOverFlowAmt As Double
  Dim DedOverFlowYTD As Double
  Dim MaskType$
  Dim TempEmp$
  Dim Nexty As Integer
  Dim OpenHandle As Integer
  Dim OK As String
  Dim SysRec As RegDSysFileRecType
  Dim SysHandle As Integer
  Dim TempHandle As Integer
  Dim ARName$
  Dim ARHandle As Integer
  Dim dlm$
  Dim SAmt As Double, HAmt#
  Dim VAmt As Double, PAmt#
  Dim LeaveHandle As Integer
  Dim NumLeaveRec As Integer
  ReDim LeaveRec(1) As LeaveRecType
  Dim CityStateZip$
  Dim BankDraft As Boolean
  Dim SSN As Boolean, y As Integer
  Dim Employer As String
  
  SSN = True
  BankDraft = True
  
  OpenLeaveFileName LeaveHandle
  NumLeaveRec = LOF(LeaveHandle) \ Len(LeaveRec(1))
  
  dlm$ = "~"
  ARName = "PRRPTS\Errorchk"
  ARHandle = FreeFile
'  On Error GoTo ErrorHandler
  Open ARName For Output As ARHandle
  Close ARHandle
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec
  Close SysHandle
  CheckStyle = SysRec.CheckStyle
  
  FF$ = Chr$(12)
  
  OK = "OK"
  OpenHandle = FreeFile
  Open "prdata\ChecksPrinted.opn" For Output As OpenHandle
  Print #OpenHandle, OK
  Close OpenHandle
  
  FreqPay$(1) = "Weekly          "
  FreqPay$(2) = "Bi-Weekly       "
  FreqPay$(3) = "Semi-Monthly    "
  FreqPay$(4) = "Monthly         "
  FreqPay$(5) = "Quarterly       "
  FreqPay$(6) = "Semi-Annually   "
  FreqPay$(7) = "Annually        "
  
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  Employer = QPTrim$(Unit(1).UFEMPR) 'added 3/15/07
  
  If Unit(1).BankDraft = "N" Then
    BankDraft = False
  End If
  
  If Unit(1).SSNOnCheck = "N" Then
    SSN = False
  End If
  
  FF$ = Chr$(12)

  If FirstEmp = 0 Then FirstEmp = 1

  Title$ = "Updating Check Information"
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show ' , Me

  TCheckNum = StartChkNum&
  ReDim Check(1) As PRCheckRecType

  ReDim PPDFInfo(1) As PRPPDraftInfoType

  ReDim DedCodes(1 To 50) As DedCodeRecType
  OpenDedCodeFile DedHandle
  For x = 1 To 50
    Get DedHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCnt = DedCnt + 1
    End If
    DedCodes(x) = DedRec
  Next x
  Close DedHandle

  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  OpenErnCodeFile ErnHandle
  For x = 1 To 3
    Get ErnHandle, x, ErnCodes(x)
    
  Next x
  Close ErnHandle
  Dim ThisErnCode$
  ThisErnCode$ = ErnCodes(1).ERNCODE1
  ThisErnCode$ = ErnCodes(2).ERNCODE1
  ThisErnCode$ = ErnCodes(3).ERNCODE1
  ReDim TransRec(1) As TransRecType
  ReDim EmpRec2(1) As EmpData2Type
  ReDim EmpRec3(1) As EmpData3Type

  TransRecLen = Len(TransRec(1))
  Emp2RecLen = Len(EmpRec2(1))
  Emp3RecLen = Len(EmpRec3(1))
  CRecLen = Len(Check(1))
  CheckRecLen = Len(Check(1))
  PPDFLen = Len(PPDFInfo(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(EmpIdxNName)
  
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) / IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  OpenChecksFile CHandle
  
  OpenEmpData2File EHandle2
  
  OpenEmpData3File EHandle3
  
  OpenTransWorkFile THandle
  
  Select Case CheckStyle
  Case 1 To 4, 7:
    RHandle = FreeFile
    KillFile "PRDATA\CHKPRNT.DAT"
    Open "PRDATA\" & CheckPrintFile For Output As RHandle
    Close RHandle
    CPFNHandle = FreeFile
    Open "PRDATA\" & CheckPrintFile For Append As #CPFNHandle
  Case 5 To 6:
    ARName = "PRRPTS\MIDCHECK.RPT"
    RHandle = FreeFile
    KillFile "PRRPTS\MIDCHECK.RPT"
    Open ARName For Output As RHandle
    Close RHandle
    CPFNHandle = FreeFile
    Open ARName For Append As #CPFNHandle
  Case Else:
    MsgBox "Please go to the System Interface Screen and select a Check Type"
    Close
    Exit Sub
  End Select
  
  
  'Employe DD Payperiod Info
  'Don't kill file if reprintflag is true
  If ReprintFlag = False Then
    If Exist("PRDATA\PPDFINFO.DAT") Then KillFile "PRDATA\PPDFINFO.DAT"
  End If
  OpenPPDraftInfo PPDFFile
  Nexty = 1
  For cnt = FirstEmp To NumOfRecs
    VAmt = 0
    SAmt = 0
    HAmt = 0
    PAmt = 0
    If Num2Print > 0 Then
      If DoneCnt = Num2Print Then
        GoTo DoneEM
      End If
    End If
    GRecNum& = CLng(IdxBuff(cnt))
    Get THandle, GRecNum&, TransRec(1)
    If TransRec(1).TActive = True And TransRec(1).NetPay > 0 Then
       Get EHandle2, GRecNum&, EmpRec2(1)
       Get EHandle3, GRecNum&, EmpRec3(1)
       'If statement allows draft info to remain intact if reprint is true
       If NumLeaveRec = 0 Then GoTo NoBenefits '01/06/2003
       Call GetBenefits(SAmt, VAmt, HAmt, PAmt, EmpRec2(1))
       
NoBenefits:
       If ReprintFlag = False Then
         If (EmpRec2(1).DRAFTCOD = "C" Or EmpRec2(1).DRAFTCOD = "S") And EmpRec2(1).PRENOTED = "Y" And BankDraft = True Then
           Check(1).DDFlag = True
           PPDFInfo(1).EmpRec = GRecNum&
           PPDFInfo(1).DraftDate = CheckDate
           PPDFInfo(1).NetPay = TransRec(1).NetPay
           Put #PPDFFile, Nexty, PPDFInfo(1)
           Nexty = Nexty + 1
         Else
           Check(1).DDFlag = False
         End If
       Else
         If (EmpRec2(1).DRAFTCOD = "C" Or EmpRec2(1).DRAFTCOD = "S") And EmpRec2(1).PRENOTED = "Y" And BankDraft = True Then
           Check(1).DDFlag = True
         Else
           Check(1).DDFlag = False
         End If
       End If
       Check(1).CActive = True
       Check(1).EmpName = QPTrim$(EmpRec2(1).EMPFNAME) + " " + QPTrim$(EmpRec2(1).EMPLNAME)
       Check(1).EmpNo = EmpRec2(1).EmpNo
       
       If SSN = True Then
         Check(1).EmpSSN = Left$(EmpRec2(1).EmpSSN, 3) + "-" + Mid$(EmpRec2(1).EmpSSN, 4, 2) + "-" + Mid$(EmpRec2(1).EmpSSN, 6, 4)
       Else
         Check(1).EmpSSN = ""
       End If
       
       Check(1).EmpAddr1 = EmpRec2(1).EmpAddr1
       Check(1).EmpCity = EmpRec2(1).EmpCity
       Check(1).EmpState = EmpRec2(1).EmpState
       Check(1).EmpZip = EmpRec2(1).EmpZip

       Check(1).PayEndDate = TransRec(1).PayPdEnd

       Check(1).CheckDate = CheckDate
       Check(1).BaseRate = TransRec(1).BaseRate
       Check(1).GrossPay = TransRec(1).GrossPay
       Check(1).FedTaxAmt = TransRec(1).FedTaxAmt
       Check(1).StaTaxAmt = TransRec(1).StaTaxAmt
       Check(1).MedTaxAmt = TransRec(1).MedTaxAmt
       Check(1).SocTaxAmt = TransRec(1).SocTaxAmt
       Check(1).TotDedAmt = TransRec(1).TotDedAmt
       Check(1).RetireAmt = TransRec(1).RetireAmt
       Check(1).TaxFring = TransRec(1).TaxFring
       Check(1).EICAmt = TransRec(1).EICAmt

       For Cnt2 = 1 To 3
         Check(1).AEarn(Cnt2).DAmt = TransRec(1).EAmt(Cnt2)
         Check(1).AEarn(Cnt2).DCode = ErnCodes(Cnt2).ERNCODE1
       Next

       Check(1).TotAdditEarn = TransRec(1).TotAdditEarn
       Check(1).NetPay = TransRec(1).NetPay

'02-16-95 ?????????
       Check(1).TotOTWage = TransRec(1).TotOTWage
'was NEVER printed on the CHECK correctly!!!!!!!!

       Check(1).TotRegWage = TransRec(1).TotRegWage
       Check(1).RegHrsWork = TransRec(1).RegHrsWork
       Check(1).RegHrsPaid = TransRec(1).RegHrsPaid
       Check(1).OTHrsPaid = TransRec(1).OTHrsPaid

       Check(1).YTDGrossPay = OldRound#(EmpRec3(1).YTDGrossPay + TransRec(1).GrossPay)
       Check(1).YTDFederal = OldRound#(EmpRec3(1).YTDFederal + TransRec(1).FedTaxAmt)
       Check(1).YTDState = OldRound#(EmpRec3(1).YTDState + TransRec(1).StaTaxAmt)
       Check(1).YTDSocial = OldRound#(EmpRec3(1).YTDSocial + TransRec(1).SocTaxAmt)
       Check(1).YTDMedicare = OldRound#(EmpRec3(1).YTDMedicare + TransRec(1).MedTaxAmt)
       Check(1).YTDTotDed = OldRound#(EmpRec3(1).YTDDAmtT + EmpRec3(1).YTDRetire + TransRec(1).TotDedAmt)
       Check(1).YTDNetPay = OldRound#(EmpRec3(1).YTDNet + TransRec(1).NetPay)
       Check(1).YTDRetire = OldRound#(EmpRec3(1).YTDRetire + TransRec(1).RetireAmt)

       Check(1).VactBal = OldRound#(EmpRec2(1).EMPVBAL - TransRec(1).VacUsed)
       If Check(1).VactBal < -20 Then Check(1).VactBal = 0

       Check(1).SickBal = OldRound#(EmpRec2(1).EMPSLBAL - TransRec(1).SickUsed)
       If Check(1).SickBal < -20 Then Check(1).SickBal = 0

'       Check(1). = OldRound#(EmpRec2(1).EMPSLBAL - TransRec(1).SickUsed)
       If Check(1).SickBal < -20 Then Check(1).SickBal = 0

       Check(1).CompEarn = OldRound#(EmpRec2(1).EMPCTE + TransRec(1).OT2Comp)
       Check(1).CompBal = OldRound#(Check(1).CompEarn - (EmpRec2(1).EMPCTUSE + TransRec(1).CompUsed))
       If Check(1).CompBal < -20 Then Check(1).CompBal = 0

       HOLBAL# = OldRound#(EmpRec2(1).HOLERN - (EmpRec2(1).HolUsed + TransRec(1).HOLHOURS))
       PERBAL# = OldRound#(EmpRec2(1).PERERN - (EmpRec2(1).PerUsed + TransRec(1).PerHours))
       
       Check(1).VacUsed = TransRec(1).VacUsed
       Check(1).SickUsed = TransRec(1).SickUsed
       Check(1).CompUsed = TransRec(1).CompUsed
       Check(1).HolUsed = TransRec(1).HOLHOURS
       Check(1).PerUsed = TransRec(1).PerHours
      'Check(1).PERUSED = TransRec(1).PerHours

        
       For Cnt2 = 1 To DedCnt
         Check(1).CDED(Cnt2).DCode = ""
         Check(1).CDED(Cnt2).DAmt = 0
         Check(1).CDED(Cnt2).YTDDAmt = 0
       Next
      
       Nextx = 1
       DedOverFlowAmt = 0
       DedOverFlowYTD = 0
       'checks only itemize 11 deduction...all the rest are a total
       For Cnt2 = 1 To DedCnt
         If TransRec(1).DAmt(Cnt2) > 0 Or EmpRec3(1).YTDDAmt(Cnt2) > 0 Then 'added the YTDDAmt on 9/27/04 because
         'sometimes deductions are changed in the middle of a year and the customers wanted them to appear
         'on the checks even if they were not active for the current check
           If Nextx = 12 And DedCnt > 12 Then
             Check(1).CDED(12).DCode = "Balance"
             DedOverFlowAmt = DedOverFlowAmt + TransRec(1).DAmt(Cnt2)
             DedOverFlowYTD = DedOverFlowYTD + OldRound#(EmpRec3(1).YTDDAmt(Cnt2) + TransRec(1).DAmt(Cnt2))
             Check(1).CDED(12).DAmt = DedOverFlowAmt
             Check(1).CDED(12).YTDDAmt = DedOverFlowYTD
             GoTo NextxMax
           End If
           Check(1).CDED(Nextx).DCode = DedCodes(Cnt2).DCDESC1
           Check(1).CDED(Nextx).DAmt = TransRec(1).DAmt(Cnt2)
           Check(1).CDED(Nextx).YTDDAmt = OldRound#(EmpRec3(1).YTDDAmt(Cnt2) + TransRec(1).DAmt(Cnt2))
           Nextx = Nextx + 1
         End If
NextxMax:
       Next
      
       For Cnt2 = 1 To DedCnt
         If Check(1).CDED(Cnt2).DAmt = 0 And Check(1).CDED(Cnt2).YTDDAmt = 0 Then 'added the YTDDAmt on 9/27/04
           Check(1).CDED(Cnt2).DCode = "" '"Unused" '8/7
         End If
       Next

       Check(1).CheckDate = CheckDate
       Check(1).CheckNum = TCheckNum&
      'TCheckNum& = TCheckNum& + 1

'*NEW 01-16-96  Added calc here for (Vac, Sick, Hol, and Comp pay)

       PayType$ = Left$(UCase$(QPTrim$(EmpRec2(1).EMPPTYPE)), 1)

       Select Case PayType$
         Case "S"
           For FreqCnt = 1 To 7
             If UCase(EmpRec2(1).EMPPFREQ) = UCase(FreqPay$(FreqCnt)) Then
               Exit For
             End If
           Next
           Select Case FreqCnt
             Case 1
               PayFreq = 52
             Case 2
               PayFreq = 26
             Case 3
               PayFreq = 24
             Case 4
               PayFreq = 12
             Case 5
               PayFreq = 4
             Case 6
               PayFreq = 2
             Case 7
               PayFreq = 1
           End Select
           YGross# = OldRound#(EmpRec2(1).EMPPRATE * PayFreq)
           HRate# = OldRound#(YGross# / 2080)
           WorkPay# = Check(1).BaseRate
        Case "H"
           HRate# = EmpRec2(1).EMPPRATE
           WorkPay# = OldRound#(Check(1).BaseRate * Check(1).RegHrsWork)
      End Select
      If TransRec(1).VacUsed > 0 Then
        VacPay# = OldRound#(HRate# * TransRec(1).VacUsed)
      Else
        VacPay# = 0
      End If

      If TransRec(1).SickUsed > 0 Then
        SickPay# = OldRound#(HRate# * TransRec(1).SickUsed)
      Else
        SickPay# = 0
      End If

      If TransRec(1).CompUsed > 0 Then
        CompPay# = OldRound#(HRate# * TransRec(1).CompUsed)
      Else
        CompPay# = 0
      End If

      If TransRec(1).HOLHOURS > 0 Then
        HolPay# = OldRound#(HRate# * TransRec(1).HOLHOURS)
      Else
        HolPay# = 0
      End If
     
      If TransRec(1).PerHours > 0 Then
        PerPay# = OldRound#(HRate# * TransRec(1).PerHours)
      Else
        PerPay# = 0
      End If
     
     Select Case CheckStyle
     Case 1:
       GoSub P901339
     Case 2:
       GoSub P901342
     Case 3:
       GoSub P9028
     Case 4:
       GoSub P9007
     Case 5:
'       GoSub LaserMidPrint
       GoSub LaserMidPrintNR
     Case 6:
       GoSub LaserTopPrint
     Case 7:
       Call Print42LineCommon(CPFNHandle, Check(), WorkPay#, HolPay#, SickPay#, CompPay#, VacPay#, TCheckNum&, CheckDate)
     Case Else:
       MsgBox "Please go to the System Interface Screen and select a Check Type"
       Close
       Exit Sub
     End Select
     
     TCheckNum& = TCheckNum& + 1
       If Num2Print > 0 Then
         DoneCnt = DoneCnt + 1
       End If

    Else
      Check(1).CActive = False
      Check(1).CheckNum = 0 'added 7/8/02
      Check(1).CheckDate = CheckDate
    End If
    
    Put CHandle, GRecNum&, Check(1)
    
DoneEM:
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next
'**UNREM
  Select Case CheckStyle
    Case 1 To 4, 7
      RPTSetupPRN 123, CPFNHandle
    Case 5 To 6
    Case Else
  End Select
'  RPTSetupPRN 123, CPFNHandle

  Close ARHandle
  Close CPFNHandle
  Close PPDFFile
  Close EHandle2
  Close EHandle3
  Close THandle
  Close CHandle
  
  Close
 
  If CheckStyle = 1 Then
    MaskType = "PRData\P9013-39MSK.txt"
  ElseIf CheckStyle = 2 Then
    MaskType = "PRData\P9013-42MSK.txt"
  ElseIf CheckStyle = 3 Then
    MaskType = "PRData\P9028MSK.txt"
  ElseIf CheckStyle = 4 Then
    MaskType = "PRData\P9007MSK.txt"
  ElseIf CheckStyle = 5 Then
    MaskType = "PRData\Laser1MSK.txt"
  Else
    MaskType = ""
  End If
    
  Select Case CheckStyle
    Case 1 To 4, 7
      ViewPrint "PRDATA\" & CheckPrintFile, "", True, True, doAlign, MaskType
      doAlign = False
    Case 5 To 6
      arMidPrint.Show
  End Select
  
  Exit Sub

P9007:
'--This is the "old standard" payroll check. Product 9007
     RPTSetupPRN 15, CPFNHandle '7/20
     Print #CPFNHandle, "~"
     Print #CPFNHandle, "Chk #" + Using("#######", TCheckNum&)

     '--Line 3
     Print #CPFNHandle, Check(1).EmpName;
     Print #CPFNHandle, Tab(65); "Rate:", Using("$###0.00", Check(1).BaseRate)

     Print #CPFNHandle,

     '--Line 5 Desc Only
     'PRINT #1,
     Print #CPFNHandle, Tab(75); "Other 1";
     Print #CPFNHandle, Tab(88); "Other 2"

     '--Line 6 Hours Section
     Print #CPFNHandle, Tab(2); Left$(QPTrim$(Check(1).EmpNo), 5);
     Print #CPFNHandle, Tab(8); Using("##0.00", Check(1).RegHrsWork);
     Print #CPFNHandle, Tab(16); Using("##0.00", Check(1).OTHrsPaid);
     Print #CPFNHandle, Tab(24); Using("##0.00", Check(1).HolUsed);
     Print #CPFNHandle, Tab(32); Using("##0.00", Check(1).CompUsed);
     '--Earnings section
     Print #CPFNHandle, Tab(42); Using("###0.00", WorkPay#);
     Print #CPFNHandle, Tab(52); Using("###0.00", VacPay#);
     Print #CPFNHandle, Tab(63); Using("###0.00", SickPay#);
     Print #CPFNHandle, Tab(73); Using("###0.00", Check(1).AEarn(1).DAmt);
     Print #CPFNHandle, Tab(88); Using("###0.00", Check(1).AEarn(2).DAmt)

     '--Line 7
     Print #CPFNHandle, Tab(75); "Other 3"

     '--Line 8 Hours
     Print #CPFNHandle, Tab(2); Check(1).EmpSSN;
     Print #CPFNHandle, Tab(16); Using("##0.00", Check(1).VacUsed);
     Print #CPFNHandle, Tab(24); Using("##0.00", Check(1).SickUsed);
     Print #CPFNHandle, Tab(32); Using("##0.00", AllHrsPaid#);

     '--Line 8 Earnings
     Print #CPFNHandle, Tab(42); Using("###0.00", Check(1).TotOTWage);
     Print #CPFNHandle, Tab(52); Using("###0.00", HolPay#);
     Print #CPFNHandle, Tab(63); Using("###0.00", CompPay#);
     Print #CPFNHandle, Tab(73); Using("###0.00", Check(1).AEarn(3).DAmt);
     Print #CPFNHandle, Tab(88); Using("###0.00", TransRec(1).GrossPay)

     '--Line 9
     Print #CPFNHandle, Tab(88); "Adv EIC"
     Print #CPFNHandle, Tab(88); Using("###0.00", Check(1).EICAmt)

     '--Line 11
     Print #CPFNHandle, Tab(31); QPTrim$(Left$(Check(1).CDED(1).DCode, 6));
     Print #CPFNHandle, Tab(38); QPTrim$(Left$(Check(1).CDED(2).DCode, 6));
     Print #CPFNHandle, Tab(47); QPTrim$(Left$(Check(1).CDED(3).DCode, 6));
     Print #CPFNHandle, Tab(57); QPTrim$(Left$(Check(1).CDED(4).DCode, 6));
     Print #CPFNHandle, Tab(67); QPTrim$(Left$(Check(1).CDED(5).DCode, 6));
     Print #CPFNHandle, Tab(77); QPTrim$(Left$(Check(1).CDED(6).DCode, 6))

     '--Line 12
     Print #CPFNHandle, Tab(2); Using("###0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));

     Print #CPFNHandle, Tab(11); Using("###0.00", Check(1).StaTaxAmt);
     Print #CPFNHandle, Tab(29); Using("###0.00", Check(1).CDED(1).DAmt);
     Print #CPFNHandle, Tab(38); Using("###0.00", Check(1).CDED(2).DAmt);
     Print #CPFNHandle, Tab(46); Using("###0.00", Check(1).CDED(3).DAmt);
     Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(4).DAmt);
     Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(5).DAmt);
     Print #CPFNHandle, Tab(76); Using("###0.00", Check(1).CDED(6).DAmt);
     Print #CPFNHandle, Tab(88); Using("###0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt))  'TotTaxes#

     '--Line 13
     Print #CPFNHandle, Tab(12); "Retire";
     Print #CPFNHandle, Tab(31); QPTrim$(Left$(Check(1).CDED(7).DCode, 6));
     Print #CPFNHandle, Tab(38); QPTrim$(Left$(Check(1).CDED(8).DCode, 6));
     Print #CPFNHandle, Tab(47); QPTrim$(Left$(Check(1).CDED(9).DCode, 6));
     Print #CPFNHandle, Tab(57); QPTrim$(Left$(Check(1).CDED(10).DCode, 6));
     Print #CPFNHandle, Tab(67); QPTrim$(Left$(Check(1).CDED(11).DCode, 6));
     Print #CPFNHandle, Tab(77); QPTrim$(Left$(Check(1).CDED(12).DCode, 6))

     '--Line 14
     Print #CPFNHandle, Tab(2); Using("###0.00", Check(1).FedTaxAmt);
     Print #CPFNHandle, Tab(11); Using("###0.00", Check(1).RetireAmt);
     Print #CPFNHandle, Tab(29); Using("###0.00", Check(1).CDED(7).DAmt);
     Print #CPFNHandle, Tab(38); Using("###0.00", Check(1).CDED(8).DAmt);
     Print #CPFNHandle, Tab(46); Using("###0.00", Check(1).CDED(9).DAmt);
     Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(10).DAmt);
     Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(11).DAmt);
     Print #CPFNHandle, Tab(76); Using("###0.00", Check(1).CDED(12).DAmt);
     Print #CPFNHandle, Tab(88); Using("###0.00", Check(1).TotDedAmt)   'TotVolDed#

     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle,

     '--Line 18
     Print #CPFNHandle, Tab(2); Using("#####0.00", Check(1).YTDGrossPay);
     Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
     Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDFederal);
     Print #CPFNHandle, Tab(36); Using("####0.00", Check(1).YTDState);
     Print #CPFNHandle, Tab(49); Using("###0.00", Check(1).YTDRetire);
     Print #CPFNHandle, Tab(71); MakeRegDate(Check(1).PayEndDate);
     Print #CPFNHandle, Tab(85); Using("####0.00", Check(1).NetPay)

     Print #CPFNHandle,

     Print #CPFNHandle, Tab(3); "Unused Vaca"; Tab(14); Using("#,##0.00", Check(1).VactBal); '8/15 respaced
     Print #CPFNHandle, Tab(25); "Unused Sick"; Tab(37); Using("#,##0.00", Check(1).SickBal); '8/15 respaced
     Print #CPFNHandle, Tab(48); "Unused Comp"; Tab(61); Using("#,##0.00", Check(1).CompBal); '8/15 respaced
     Print #CPFNHandle, Tab(73); "Unused Hl/Pr"; Tab(87); Using("#,##0.00", HOLBAL# + PERBAL#)
     '--Line 21 - Last line of stub"
     Print #CPFNHandle, Tab(3); "Other Taxable"; Tab(17); Using("$##,##0.00", Check(1).TaxFring)

     If Check(1).DDFlag = True Then
       Print #CPFNHandle,
       Print #CPFNHandle, "   DIRECT DEPOSIT VOUCHER"


       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(12); Check(1).EmpName
       Print #CPFNHandle, Tab(12); Check(1).EmpAddr1
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); ", "; Check(1).EmpState; " "; Check(1).EmpZip

       Print #CPFNHandle,
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID            VOID"
       Print #CPFNHandle, "_"

     Else
       Print #CPFNHandle, '--Line 22 - Top of check
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(71); MakeRegDate(Check(1).CheckDate)
       Print #CPFNHandle, 'TAB(73); Check(1).CheckNum
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(77); Using("$####0.00", Check(1).NetPay)
       Print #CPFNHandle, Tab(5); SpellNumber$(Using$("####0.00", Check(1).NetPay))
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(12); Check(1).EmpName
       Print #CPFNHandle, Tab(12); Check(1).EmpAddr1
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); ", "; Check(1).EmpState; " "; Check(1).EmpZip
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, "~"
'       Print #CPFNHandle, 'FF$ 7/22
     End If
  Return

'39 line "new standard" check. product 9013
P901339:
    '--Line 1
    RPTSetupPRN 15, CPFNHandle
    Print #CPFNHandle, "~"
'    Print #CPFNHandle,
    If InStr(UCase(Employer), "CREWE") Then
      Print #CPFNHandle,
    Else
      Print #CPFNHandle, Tab(20); Left$(QPTrim$(Check(1).EmpNo) + " " + QPTrim$(Check(1).EmpName), 26); Tab(48); "PPE: "; MakeRegDate(Check(1).PayEndDate); Tab(64); "Rate:"; Tab(70); Using("###0.00", Check(1).BaseRate)
    End If

    '--Line 2
    If Check(1).DDFlag = True Then
      Print #CPFNHandle, "   DIRECT DEPOSIT VOUCHER"
    Else
      Print #CPFNHandle,
    End If

    '--Line 3
    Print #CPFNHandle, "                HRS     EARN        YTD             PERIOD           YTD"

    '--Line 4
    Print #CPFNHandle, Tab(1); "HRS WORKED"; Tab(13); Using("###0.00", Check(1).RegHrsWork);
    Print #CPFNHandle, Tab(21); Using("####0.00", WorkPay#);
'    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).RegHrsPaid)
    Print #CPFNHandle, Tab(41); "RETIREMENT";
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).RetireAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).YTDRetire);

    '--Line 5
    Print #CPFNHandle, Tab(1); "HOL/PERS HRS"; Tab(13); Using("###0.00", (Check(1).HolUsed + Check(1).PerUsed));
    Print #CPFNHandle, Tab(21); Using("####0.00", (HolPay# + PerPay#));
    Print #CPFNHandle, Tab(32); Using("####0.00", (HOLBAL# + PERBAL#));
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(1).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(1).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(1).YTDDAmt)

    '--Line 6
    Print #CPFNHandle, Tab(1); "SICK HRS"; Tab(13); Using("###0.00", Check(1).SickUsed);
    Print #CPFNHandle, Tab(21); Using("####0.00", SickPay#);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).SickBal);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(2).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(2).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(2).YTDDAmt)

    '--Line 7
    Print #CPFNHandle, Tab(1); "COMP HRS"; Tab(13); Using("###0.00", Check(1).CompUsed);
    Print #CPFNHandle, Tab(21); Using("####0.00", CompPay#);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).CompBal);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(3).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(3).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(3).YTDDAmt);

    '--Line 8
    Print #CPFNHandle, Tab(1); "VAC HRS"; Tab(13); Using("###0.00", Check(1).VacUsed);
    Print #CPFNHandle, Tab(21); Using("####0.00", VacPay#);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).VactBal);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(4).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(4).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(4).YTDDAmt);

    '--Line 9
    Print #CPFNHandle, Tab(1); "TOT REG HRS"; Tab(13); Using("###0.00", Check(1).RegHrsPaid);
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).TotRegWage);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(5).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(5).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(5).YTDDAmt);

    '--Line 10
    Print #CPFNHandle, Tab(1); "OT  HRS"; Tab(13); Using("###0.00", Check(1).OTHrsPaid);
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).TotOTWage);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(6).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(6).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(6).YTDDAmt);

    '--Line 11
    Print #CPFNHandle, Tab(1); QPTrim$(Check(1).AEarn(1).DCode);
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).AEarn(1).DAmt);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(7).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(7).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(7).YTDDAmt)

    '--Line 12
    Print #CPFNHandle, Tab(1); QPTrim(Check(1).AEarn(2).DCode);
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).AEarn(2).DAmt);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(8).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(8).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(8).YTDDAmt)

    '--Line 13
    Print #CPFNHandle, Tab(1); QPTrim$(Check(1).AEarn(3).DCode);
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).AEarn(3).DAmt);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(9).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(9).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(9).YTDDAmt)

    '--Line 14
    Print #CPFNHandle, Tab(1); "GROSS PAY";
    Print #CPFNHandle, Tab(21); Using("####0.00", Check(1).GrossPay);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).YTDGrossPay);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(10).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(10).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(10).YTDDAmt)

   '--Line 15
    Print #CPFNHandle, "FED TAX"; Tab(21); Using("####0.00", Check(1).FedTaxAmt);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).YTDFederal);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(11).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(11).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(11).YTDDAmt)

    '--Line 16
    Print #CPFNHandle, "STA TAX"; Tab(21); Using("####0.00", Check(1).StaTaxAmt);
    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).YTDState);
    Print #CPFNHandle, Tab(41); QPTrim$(Check(1).CDED(12).DCode);
    Print #CPFNHandle, Tab(52); Using("###0.00", Check(1).CDED(12).DAmt);
    Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(12).YTDDAmt)

    '--Line 17
    Print #CPFNHandle, "FICA"; Tab(21); Using("####0.00", OldRound(Check(1).MedTaxAmt + Check(1).SocTaxAmt));

    Print #CPFNHandle, Tab(32); Using("####0.00", Check(1).YTDSocial + Check(1).YTDMedicare);
    Print #CPFNHandle, Tab(41); "NET PAY";
    Print #CPFNHandle, Tab(51); Using("####0.00", Check(1).NetPay);
    Print #CPFNHandle, Tab(65); Using("####0.00", Check(1).YTDNetPay)

    '--Line 18 - Last line of stub
    Print #CPFNHandle, "_"

    If Check(1).DDFlag = True Then
    '--Line 19 - First line of check
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"; Tab(70); Using("#######", TCheckNum&)
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(3); SpellNumber$(Using$("####0.00", Check(1).NetPay)) '--Line 30
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(53); MakeRegDate(CheckDate); Tab(68); Using("$##,##0.00", Check(1).NetPay)
      Print #CPFNHandle, Tab(10); Left$(QPTrim$(Check(1).EmpName), 45)
      Print #CPFNHandle, Tab(10); Tab(10); QPTrim$(Check(1).EmpAddr1)
      Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
      Print #CPFNHandle, "_" '--Line 39

    Else
    '--Line 19 - First line of check
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(70); Using("#######", TCheckNum&)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(3); SpellNumber$(Using$("####0.00", Check(1).NetPay)) '--Line 30
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(53); MakeRegDate(CheckDate); Tab(68); Using("$##,##0.00", Check(1).NetPay)
      Print #CPFNHandle, Tab(10); Left$(QPTrim$(Check(1).EmpName), 45)
      Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpAddr1)
      Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, "_" '--Line 39
    End If
  Return

P901342:
'This is the "CARTHAGE" check format 3/11/96
     RPTSetupPRN 15, CPFNHandle '7/20
'1
     Print #CPFNHandle, "~"; Tab(20); QPTrim$(Check(1).EmpNo); Tab(30); QPTrim$(Check(1).EmpName); '7/24 altered to include ~ and made Tab(10) to Tab(9)
     Print #CPFNHandle, Tab(54); QPTrim$(Check(1).EmpSSN);
     Print #CPFNHandle, Tab(67); "PPE: "; MakeRegDate(Check(1).PayEndDate);
     Print #CPFNHandle, 'Tab(82); " "; Using("######", TCheckNum&)
'2
     Print #CPFNHandle, "                HRS          REG           YTD    DEDUCTIONS              PER            YTD"
'3
     Print #CPFNHandle, Tab(1); "REG HRS"; Tab(13); Using("###0.00", Check(1).RegHrsWork);
     Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotRegWage);
     Print #CPFNHandle, Tab(39); Using("####0.00", OldRound#(EmpRec3(1).YTDRegPay + Check(1).TotRegWage));

     If QPTrim$(Check(1).CDED(1).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(1).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(1).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(1).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'4
     Print #CPFNHandle, Tab(1); "OT  HRS"; Tab(13); Using("###0.00", Check(1).OTHrsPaid);        'Check(1).TotOTWage;
     Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotOTWage);
'fixed
     Print #CPFNHandle, Tab(39); Using("####0.00", OldRound#(EmpRec3(1).YTDOTPay + Check(1).TotOTWage));
'here
     
     If QPTrim$(Check(1).CDED(2).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(2).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(2).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(2).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'5
     Print #CPFNHandle, "ADD EARN"; '    ###0.00";
     Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotAdditEarn);
     Print #CPFNHandle, Tab(39); Using("####0.00", OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn));

     If QPTrim$(Check(1).CDED(3).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(3).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(3).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(3).YTDDAmt) ';
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle, ';
     End If
'6
     Print #CPFNHandle, Tab(1); "VACT BAL"; Tab(13); Using("###0.00", Check(1).VacUsed);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).VactBal);
     
     If QPTrim$(Check(1).CDED(4).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(4).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(4).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(4).YTDDAmt) ';
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle, ';
     End If
'7
     Print #CPFNHandle, Tab(1); "SICK BAL"; Tab(13); Using("###0.00", Check(1).SickUsed);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).SickBal);
     
     If QPTrim$(Check(1).CDED(5).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(5).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(5).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(5).YTDDAmt) ';
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle, ';
     End If
'8
     Print #CPFNHandle, Tab(1); "COMP BAL"; Tab(13); Using("###0.00", Check(1).CompUsed);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).CompBal);
     
     If QPTrim$(Check(1).CDED(6).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(6).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(6).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(6).YTDDAmt) ';
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle, ';
     End If
'9
     Print #CPFNHandle, Tab(1); "HOL BAL"; Tab(13); Using("###0.00", Check(1).HolUsed);
     Print #CPFNHandle, Tab(39); Using("####0.00", HOLBAL#);
     
     If QPTrim$(Check(1).CDED(7).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(7).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(7).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(7).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'10
     Print #CPFNHandle, Tab(1); "PER BAL"; Tab(13); Using("###0.00", Check(1).PerUsed);
     Print #CPFNHandle, Tab(39); Using("####0.00", PERBAL#);
     
     If QPTrim$(Check(1).CDED(8).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(8).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(8).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(8).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'11
     Print #CPFNHandle, "FED TAX"; Tab(25); Using("####0.00", Check(1).FedTaxAmt);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDFederal);
     
     If QPTrim$(Check(1).CDED(9).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(9).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(9).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(9).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'12
     Print #CPFNHandle, "STA TAX"; Tab(25); Using("####0.00", Check(1).StaTaxAmt);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDState);
     
     
     If QPTrim$(Check(1).CDED(10).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(10).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(10).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(10).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'13
     Print #CPFNHandle, "FICA"; Tab(25); Using("####0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));

     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDSocial + Check(1).YTDMedicare);

     If QPTrim$(Check(1).CDED(11).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(11).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(11).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(11).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'14
     Print #CPFNHandle, "RET "; Tab(25); Using("####0.00", Check(1).RetireAmt);
     Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDRetire);

     If QPTrim$(Check(1).CDED(12).DCode) <> "" Then
       Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(12).DCode);
       Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(12).DAmt);
       Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(12).YTDDAmt)
     Else
       Print #CPFNHandle, ;
       Print #CPFNHandle, ;
       Print #CPFNHandle,
     End If
'15
     Print #CPFNHandle, "   RATE       GROSS        TOT DED           NET       YTDGROSS     YTD DED          YTD NET"

     Print #CPFNHandle, Using("###0.00", Check(1).BaseRate);
     Print #CPFNHandle, Tab(12); Using("####0.00", Check(1).GrossPay);
     Print #CPFNHandle, Tab(28); Using("###0.00", OldRound(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).TotDedAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt));

     Print #CPFNHandle, Tab(41); Using("####0.00", Check(1).NetPay);
     Print #CPFNHandle, Tab(56); Using("####0.00", Check(1).YTDGrossPay);
     Print #CPFNHandle, Tab(68); Using("####0.00", OldRound(Check(1).YTDTotDed + Check(1).YTDFederal + Check(1).YTDState + Check(1).YTDSocial + Check(1).YTDMedicare));
     Print #CPFNHandle, Tab(85); Using("####0.00", Check(1).YTDNetPay)

'End of Stub
'     For Cnt2 = 1 To 4
'16,17,18,19,20,21
     For Cnt2 = 1 To 6
       Print #CPFNHandle, '"***"
     Next
     
     If Check(1).DDFlag = True Then
       Print #CPFNHandle, Tab(83); Using("#######", TCheckNum&)
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); SpellNumber$(Using$("####0.00", Check(1).NetPay)) 'Print the whole number part
      
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(65); MakeRegDate(CheckDate); Tab(79); Using("$##,##0.00", Check(1).NetPay)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpName)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpAddr1)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, ""
       Print #CPFNHandle, ""
       Print #CPFNHandle, ""
       Print #CPFNHandle, ""
       Print #CPFNHandle, "~" 'added ~ on 7/24
     Else
 'start @ 22
       Print #CPFNHandle, Tab(83); Using("#######", TCheckNum&)
       Print #CPFNHandle, '"24"
       Print #CPFNHandle, '"25"
       Print #CPFNHandle, '"26"
       Print #CPFNHandle, '"27"
       Print #CPFNHandle, Tab(10); SpellNumber$(Using$("####0.00", Check(1).NetPay)) 'Print the whole number part
      
       Print #CPFNHandle, '"29"
       Print #CPFNHandle, '"30"
       Print #CPFNHandle, Tab(65); MakeRegDate(CheckDate); Tab(79); Using("$##,##0.00", Check(1).NetPay)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpName)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpAddr1)
       Print #CPFNHandle, Tab(10); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
       Print #CPFNHandle, '"35"
       Print #CPFNHandle, '"36"
       Print #CPFNHandle, '"37"
       Print #CPFNHandle, '"38"
       Print #CPFNHandle, '"39"
       Print #CPFNHandle, '"40"
       Print #CPFNHandle, '"41"
'42
       Print #CPFNHandle, "~"
   End If
   Return
'*******************************************************************************************
  
P9028:
'This is the "STANDARD" check format 6/27/94
     RPTSetupPRN 15, CPFNHandle '7/20
     TempEmp$ = QPTrim$(Left$(Check(1).EmpName, 33))
     'Print #CPFNHandle, CHR$(27) + CHR$(58);'sets compressed mode   12 dpi
     'Print #CPFNHandle, "Top line                                                       Top Line"
     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle, Tab(2); TempEmp$; Tab(66); "Rate:"; Using("###0.00", Check(1).BaseRate)
     Print #CPFNHandle,
     Print #CPFNHandle, Tab(65); "Tax Frng"; Tab(76); "EIC"

     Print #CPFNHandle, Tab(1); QPTrim$(Left$(LTrim$(Check(1).EmpNo), 5));

     'IF QPTrim$(Check(1).EmpNo) = "6" THEN STOP
     Print #CPFNHandle, Tab(7); Using("##0.00", Check(1).RegHrsWork);
'-new
     Print #CPFNHandle, Tab(15); Using("##0.00", Check(1).OTHrsPaid);
     Print #CPFNHandle, Tab(22); Using("##0.00", Check(1).HolUsed);
     Print #CPFNHandle, Tab(29); Using("##0.00", Check(1).CompUsed);
     Print #CPFNHandle, Tab(37); Using("###0.00", Check(1).TotRegWage);
     Print #CPFNHandle, Tab(64); Using("####0.00", Check(1).TaxFring);
     Print #CPFNHandle, Tab(73); Using("###0.00", Check(1).EICAmt)

     Print #CPFNHandle, Tab(45); QPTrim$(Check(1).AEarn(3).DCode);
     Print #CPFNHandle, Tab(55); QPTrim$(Check(1).AEarn(2).DCode);
     Print #CPFNHandle, Tab(64); QPTrim$(Check(1).AEarn(1).DCode)

     Print #CPFNHandle, Tab(2); QPTrim$(Check(1).EmpSSN);
     Print #CPFNHandle, Tab(15); Using("##0.00", Check(1).VacUsed);
     Print #CPFNHandle, Tab(22); Using("##0.00", Check(1).SickUsed);
     Print #CPFNHandle, Tab(29); Using("##0.00", Check(1).RegHrsPaid);
     Print #CPFNHandle, Tab(37); Using("###0.00", Check(1).TotOTWage);

     Print #CPFNHandle, Tab(47); Using("###0.00", Check(1).AEarn(3).DAmt);
     Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).AEarn(2).DAmt);
     Print #CPFNHandle, Tab(65); Using("###0.00", Check(1).AEarn(1).DAmt);

     Print #CPFNHandle, Tab(73); Using("###0.00", Check(1).GrossPay)
     Print #CPFNHandle,
     Print #CPFNHandle,

     Print #CPFNHandle, Tab(28); QPTrim$(Left$(Check(1).CDED(1).DCode, 6));
     Print #CPFNHandle, Tab(35); QPTrim$(Left$(Check(1).CDED(2).DCode, 6));
     Print #CPFNHandle, Tab(43); QPTrim$(Left$(Check(1).CDED(3).DCode, 6));
     Print #CPFNHandle, Tab(51); QPTrim$(Left$(Check(1).CDED(4).DCode, 6));
     Print #CPFNHandle, Tab(59); QPTrim$(Left$(Check(1).CDED(5).DCode, 6));
     Print #CPFNHandle, Tab(67); QPTrim$(Left$(Check(1).CDED(6).DCode, 6))

     Print #CPFNHandle, Tab(2); Using("###0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));

     Print #CPFNHandle, Tab(12); Using("###0.00", Check(1).StaTaxAmt);

     Print #CPFNHandle, Tab(26); Using("###0.00", Check(1).CDED(1).DAmt);
     Print #CPFNHandle, Tab(33); Using("###0.00", Check(1).CDED(2).DAmt);
     Print #CPFNHandle, Tab(42); Using("###0.00", Check(1).CDED(3).DAmt);
     Print #CPFNHandle, Tab(50); Using("###0.00", Check(1).CDED(4).DAmt);
     Print #CPFNHandle, Tab(58); Using("###0.00", Check(1).CDED(5).DAmt);
     Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(6).DAmt);
     Print #CPFNHandle, Tab(73); Using("###0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt));

     Print #CPFNHandle, Tab(28); QPTrim$(Left$(Check(1).CDED(7).DCode, 6));
     Print #CPFNHandle, Tab(35); QPTrim$(Left$(Check(1).CDED(8).DCode, 6));
     Print #CPFNHandle, Tab(43); QPTrim$(Left$(Check(1).CDED(9).DCode, 6));
     Print #CPFNHandle, Tab(51); QPTrim$(Left$(Check(1).CDED(10).DCode, 6));
     Print #CPFNHandle, Tab(59); QPTrim$(Left$(Check(1).CDED(11).DCode, 6));
     Print #CPFNHandle, Tab(67); QPTrim$(Left$(Check(1).CDED(12).DCode, 6))

     Print #CPFNHandle, Tab(2); Using("###0.00", Check(1).FedTaxAmt);
     Print #CPFNHandle, Tab(12); Using("###0.00", Check(1).RetireAmt);

     Print #CPFNHandle, Tab(26); Using("###0.00", Check(1).CDED(7).DAmt);
     Print #CPFNHandle, Tab(33); Using("###0.00", Check(1).CDED(8).DAmt);
     Print #CPFNHandle, Tab(42); Using("###0.00", Check(1).CDED(9).DAmt);
     Print #CPFNHandle, Tab(50); Using("###0.00", Check(1).CDED(10).DAmt);
     Print #CPFNHandle, Tab(58); Using("###0.00", Check(1).CDED(11).DAmt);
     Print #CPFNHandle, Tab(66); Using("###0.00", Check(1).CDED(12).DAmt);

     Print #CPFNHandle, Tab(73); Using("###0.00", Check(1).TotDedAmt)

     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle, Tab(2); Using("#####0.00", Check(1).YTDGrossPay);
     Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
     Print #CPFNHandle, Tab(23); Using("####0.00", Check(1).YTDFederal);
     Print #CPFNHandle, Tab(33); Using("####0.00", Check(1).YTDState);
     Print #CPFNHandle, Tab(61); MakeRegDate(Check(1).CheckDate);
     Print #CPFNHandle, Tab(72); Using("####0.00", Check(1).NetPay)

     Print #CPFNHandle,
     Print #CPFNHandle, Tab(2); "Vac Due: "; Using("###0.00#", Check(1).VactBal);
     Print #CPFNHandle, Tab(20); "Sick Due: "; Using("###0.00#", Check(1).SickBal);
     Print #CPFNHandle, Tab(39); "Comp Due: "; Using("###0.00#", Check(1).CompBal);
     Print #CPFNHandle, Tab(58); "Hol/Pers Due: "; Using("###0.00#", (HOLBAL# + PERBAL#))


'     Print #CPFNHandle,
'     Print #CPFNHandle, Tab(4); "Vac. Due: "; Using("###0.00#", Check(1).VactBal);
'     Print #CPFNHandle, Tab(30); "Sick Due: "; Using("###0.00#", Check(1).SickBal);
'     Print #CPFNHandle, Tab(56); "Comp Due: "; Using("###0.00#", Check(1).CompBal)

     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle,
     Print #CPFNHandle, Tab(61); MakeRegDate(Check(1).CheckDate); " "; Check(1).CheckNum
     If Check(1).DDFlag = True Then
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(68); Using("$##,##0.00", Check(1).NetPay)
       Print #CPFNHandle, Tab(12); SpellNumber$(Using$("####0.00", Check(1).NetPay)); 'Print the whole number part
  
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName)
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
       Print #CPFNHandle,
       Print #CPFNHandle, "~" 'added 7/24
     Else
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(68); Using("$##,##0.00", Check(1).NetPay)
       Print #CPFNHandle, Tab(12); SpellNumber$(Using$("####0.00", Check(1).NetPay)); 'Print the whole number part
  
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName)
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
       Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle,
       Print #CPFNHandle, "~" 'added 7/24
    End If
  Return
  
LaserMidPrint:
      Print #CPFNHandle, "~ "; QPTrim$(Check(1).EmpName); Tab(30); Check(1).EmpSSN; Tab(50); "EMP NO."; Tab(59); QPTrim$(Check(1).EmpNo)
      Print #CPFNHandle, "BASE RATE: "; Tab(12); Using("####0.00", Check(1).BaseRate); Tab(24); "PERIOD END: " & MakeRegDate(Check(1).PayEndDate); Tab(50); "CHK DATE: " & MakeRegDate(CheckDate)
      '                                                                   x
      Print #CPFNHandle, "        HRS    PERIOD   RATE    USED BALANCE  Deductions   Period      YTD  "

      Print #CPFNHandle, "VAC";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", VacPay#);
      Print #CPFNHandle, Tab(24); Using("###0.00", Str(VAmt));
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).VactBal);
      If Len(QPTrim$(Check(1).CDED(1).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(1).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(1).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(1).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(1).YTDDAmt)
      End If

      Print #CPFNHandle, "SICK";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", SickPay#);
      Print #CPFNHandle, Tab(24); Using("###0.00", Str(SAmt));
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).SickBal);
      If Len(QPTrim$(Check(1).CDED(2).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(2).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(2).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(2).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(2).YTDDAmt)
      End If

      Print #CPFNHandle, "HOL/PER";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).HolUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", HolPay#);
      Print #CPFNHandle, Tab(24); Using("###0.00", Str(HAmt + PAmt));
      Print #CPFNHandle, Tab(32); Using("###0.00", (Check(1).HolUsed + Check(1).PerUsed));
      Print #CPFNHandle, Tab(40); Using("###0.00", (HOLBAL# + PERBAL#));
      If Len(QPTrim$(Check(1).CDED(3).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(3).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(3).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(3).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(3).YTDDAmt)
      End If

      Print #CPFNHandle, "COMP";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", CompPay#);
      Print #CPFNHandle, Tab(24); "" ' Using("###0.00", Check(1).CompEarn);
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).CompBal);
      If Len(QPTrim$(Check(1).CDED(4).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(4).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(4).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(4).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(4).YTDDAmt)
      End If
      Print #CPFNHandle, "        HRS    PERIOD        YTD   "
      Print #CPFNHandle, "REG";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).RegHrsWork);
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound(Check(1).RegHrsWork * Check(1).BaseRate));  'TAB(25); Check(1).YTDGrossPay;
      If Len(QPTrim$(Check(1).CDED(5).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(5).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(5).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(5).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(5).YTDDAmt)
      End If

      Print #CPFNHandle, "OT";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).OTHrsPaid);
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotOTWage);
      If Len(QPTrim$(Check(1).CDED(6).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(6).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(6).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(6).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(6).YTDDAmt)
      End If
      
      Print #CPFNHandle, "GROSS";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).GrossPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDGrossPay);
      If Len(QPTrim$(Check(1).CDED(7).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(7).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(7).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(7).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(7).YTDDAmt)
      End If

      Print #CPFNHandle, "FWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).FedTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDFederal);
      If Len(QPTrim$(Check(1).CDED(8).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(8).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(8).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(8).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(8).YTDDAmt)
      End If

      Print #CPFNHandle, "FICA";
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
      If Len(QPTrim$(Check(1).CDED(9).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(9).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(9).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(9).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(9).YTDDAmt)
      End If

      Print #CPFNHandle, "RET";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).RetireAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDRetire);
      If Len(QPTrim$(Check(1).CDED(10).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(10).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(10).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(10).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(10).YTDDAmt)
      End If

      Print #CPFNHandle, "NET PAY";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).NetPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDNetPay);
      If Len(QPTrim$(Check(1).CDED(11).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(11).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(11).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(11).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(11).YTDDAmt)
      End If
      Print #CPFNHandle, "SWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", TransRec(1).StaTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDState);
      If Len(QPTrim$(Check(1).CDED(12).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(12).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(12).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(12).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(12).YTDDAmt)
      End If
      Print #CPFNHandle, "TOT ADD";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotAdditEarn);
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn))
      'PRINT #CPFNHandle, ""
      
      If Check(1).DDFlag = True Then
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, ""
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(72); Using("######", Check(1).CheckNum)
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle,
        Print #CPFNHandle, SpellNumber$(Using$("####0.00", Check(1).NetPay))
        Print #CPFNHandle, Tab(50); MakeRegDate(CheckDate); Tab(67); Using("$#,###,0.00", Check(1).NetPay)
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle,
        Print #CPFNHandle, "~"
        Print #CPFNHandle,
        Print #CPFNHandle,
      Else
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(72); Using("######", Check(1).CheckNum)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, SpellNumber$(Using$("####0.00", Check(1).NetPay))
        Print #CPFNHandle, Tab(50); MakeRegDate(CheckDate); Tab(67); Using("$#,###,0.00", Check(1).NetPay)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, "~"
        Print #CPFNHandle,
        Print #CPFNHandle,
      End If
      Print #CPFNHandle, "~ "; QPTrim$(Check(1).EmpName); Tab(30); Check(1).EmpSSN; Tab(50); "EMP NO."; Tab(59); QPTrim$(Check(1).EmpNo)
      Print #CPFNHandle, "BASE RATE: "; Tab(12); Using("####0.00", Check(1).BaseRate); Tab(24); "PERIOD END: " & MakeRegDate(Check(1).PayEndDate); Tab(50); "CHK DATE: " & MakeRegDate(CheckDate)
      '                                                                   x
      Print #CPFNHandle, "        HRS    PERIOD   RATE    USED BALANCE  Deductions   Period      YTD  "

      Print #CPFNHandle, "VAC";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", VacPay#);
      Print #CPFNHandle, Tab(24); Using("###0.00", Str(VAmt));
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).VactBal);
      If Len(QPTrim$(Check(1).CDED(1).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(1).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(1).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(1).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(1).YTDDAmt)
      End If

      Print #CPFNHandle, "SICK";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", SickPay#);
      Print #CPFNHandle, Tab(24); Using("###0.00", Str(SAmt));
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).SickBal);
      If Len(QPTrim$(Check(1).CDED(2).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(2).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(2).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(2).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(2).YTDDAmt)
      End If

      Print #CPFNHandle, "HOL";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).HolUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", HolPay#);
'      Print #CPFNHandle, Tab(25); Using("###0.00", Check(1).HOLBAL)
      If Len(QPTrim$(Check(1).CDED(3).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(3).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(3).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(3).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(3).YTDDAmt)
      End If

      Print #CPFNHandle, "COMP";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", CompPay#);
      Print #CPFNHandle, Tab(24); "" ' Using("###0.00", Check(1).CompEarn);
      Print #CPFNHandle, Tab(32); Using("###0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(40); Using("###0.00", Check(1).CompBal);
      If Len(QPTrim$(Check(1).CDED(4).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(4).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(4).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(4).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(4).YTDDAmt)
      End If
      Print #CPFNHandle, "        HRS    PERIOD        YTD   "
      Print #CPFNHandle, "REG";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).RegHrsWork);
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound(Check(1).RegHrsWork * Check(1).BaseRate));  'TAB(25); Check(1).YTDGrossPay;
      If Len(QPTrim$(Check(1).CDED(5).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(5).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(5).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(5).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(5).YTDDAmt)
      End If

      Print #CPFNHandle, "OT";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).OTHrsPaid);
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotOTWage);
      If Len(QPTrim$(Check(1).CDED(6).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(6).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(6).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(6).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(6).YTDDAmt)
      End If
      
      Print #CPFNHandle, "GROSS";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).GrossPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDGrossPay);
      If Len(QPTrim$(Check(1).CDED(7).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(7).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(7).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(7).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(7).YTDDAmt)
      End If

      Print #CPFNHandle, "FWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).FedTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDFederal);
      If Len(QPTrim$(Check(1).CDED(8).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(8).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(8).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(8).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(8).YTDDAmt)
      End If

      Print #CPFNHandle, "FICA";
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
      If Len(QPTrim$(Check(1).CDED(9).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(9).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(9).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(9).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(9).YTDDAmt)
      End If

      Print #CPFNHandle, "RET";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).RetireAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDRetire);
      If Len(QPTrim$(Check(1).CDED(10).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(10).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(10).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(10).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(10).YTDDAmt)
      End If

      Print #CPFNHandle, "NET PAY";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).NetPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDNetPay);
      If Len(QPTrim$(Check(1).CDED(11).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(11).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(11).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(11).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(11).YTDDAmt)
      End If
      Print #CPFNHandle, "SWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", TransRec(1).StaTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDState);
      If Len(QPTrim$(Check(1).CDED(12).DCode)) = 0 Then
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(12).DCode);
        Print #CPFNHandle, Tab(61); "       ";
        Print #CPFNHandle, Tab(70); "       "
      Else
        Print #CPFNHandle, Tab(49); QPTrim$(Check(1).CDED(12).DCode);
        Print #CPFNHandle, Tab(61); Using("###0.00", Check(1).CDED(12).DAmt);
        Print #CPFNHandle, Tab(70); Using("###0.00", Check(1).CDED(12).YTDDAmt)
      End If
      Print #CPFNHandle, "TOT ADD";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotAdditEarn);
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn))
      Print #CPFNHandle, "~" '; Chr$(12)
'      Print #CPFNHandle,
  Return

LaserMidPrintNR: 'ar reports
      '                              0                               1                            2
      Print #CPFNHandle, QPTrim$(Check(1).EmpName); dlm; QPTrim$(Check(1).EmpSSN); dlm; QPTrim$(Check(1).EmpNo); dlm;
      '                        3                             4                                   5                                                                x
      Print #CPFNHandle, Check(1).BaseRate; dlm; MakeRegDate(Check(1).PayEndDate); dlm; MakeRegDate(CheckDate); dlm;
      '                         6                  7              8                   9
      Print #CPFNHandle, Check(1).VacUsed; dlm; VacPay#; dlm; Str(VAmt); dlm; Check(1).VacUsed; dlm;
      
      '                       10                               11                             12                              13
      Print #CPFNHandle, Check(1).VactBal; dlm; QPTrim$(Check(1).CDED(1).DCode); dlm; Check(1).CDED(1).DAmt; dlm; Check(1).CDED(1).YTDDAmt; dlm;

      '                         14                  15              16                17                       18
      Print #CPFNHandle, Check(1).SickUsed; dlm; SickPay#; dlm; Str(SAmt); dlm; Check(1).SickUsed; dlm; Check(1).SickBal; dlm;
      
      
      '                              19                                  20                          21
      Print #CPFNHandle, QPTrim$(Check(1).CDED(2).DCode); dlm; Check(1).CDED(2).DAmt; dlm; Check(1).CDED(2).YTDDAmt; dlm;

      '                        22                 23                         24                               25                          26
      Print #CPFNHandle, Check(1).HolUsed; dlm; HolPay#; dlm; QPTrim$(Check(1).CDED(3).DCode); dlm; Check(1).CDED(3).DAmt; dlm; Check(1).CDED(3).YTDDAmt; dlm;

      '                                                 27                                                               28                    29                         30                                            31
      Print #CPFNHandle, TransRec(1).OT2Comp; dlm; CompPay#; dlm; ""; dlm; Check(1).CompUsed; dlm; Check(1).CompBal; dlm;
     
     '                         27                  28                 29                                30                                     31
'      Print #CPFNHandle, Check(1).CompUsed; dlm; CompPay#; dlm; ""; dlm; Check(1).CompUsed; dlm; Check(1).CompBal; dlm; 'changed 2/2/2011
      
      '                                32                                 33                            34
      Print #CPFNHandle, QPTrim$(Check(1).CDED(4).DCode); dlm; Check(1).CDED(4).DAmt; dlm; Check(1).CDED(4).YTDDAmt; dlm;
      
      '                           35                                            36
      Print #CPFNHandle, Check(1).RegHrsWork; dlm; OldRound(Check(1).RegHrsWork * Check(1).BaseRate); dlm;
      
      '                               37                               38                             39
      Print #CPFNHandle, QPTrim$(Check(1).CDED(5).DCode); dlm; Check(1).CDED(5).DAmt; dlm; Check(1).CDED(5).YTDDAmt; dlm;

      '                          40                       41
      Print #CPFNHandle, Check(1).OTHrsPaid; dlm; Check(1).TotOTWage; dlm;
      
      '                                  42                             43                           44
      Print #CPFNHandle, QPTrim$(Check(1).CDED(6).DCode); dlm; Check(1).CDED(6).DAmt; dlm; Check(1).CDED(6).YTDDAmt; dlm;
      
      '                         45                        46
      Print #CPFNHandle, Check(1).GrossPay; dlm; Check(1).YTDGrossPay; dlm;
      
      '                             47                                 48                             49
      Print #CPFNHandle, QPTrim$(Check(1).CDED(7).DCode); dlm; Check(1).CDED(7).DAmt; dlm; Check(1).CDED(7).YTDDAmt; dlm;
      

      '                         50                         51
      Print #CPFNHandle, Check(1).FedTaxAmt; dlm; Check(1).YTDFederal; dlm;
      
      '                                 52                            53                            54
      Print #CPFNHandle, QPTrim$(Check(1).CDED(8).DCode); dlm; Check(1).CDED(8).DAmt; dlm; Check(1).CDED(8).YTDDAmt; dlm;
      

      '                                          55                                                            56
      Print #CPFNHandle, OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt); dlm; OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare); dlm;
      
      '                                  57                              58                               59
      Print #CPFNHandle, QPTrim$(Check(1).CDED(9).DCode); dlm; Check(1).CDED(9).DAmt; dlm; Check(1).CDED(9).YTDDAmt; dlm;
      

      '                           60                       61
      Print #CPFNHandle, Check(1).RetireAmt; dlm; Check(1).YTDRetire; dlm;
      
      '                                62                              63                              64
      Print #CPFNHandle, QPTrim$(Check(1).CDED(10).DCode); dlm; Check(1).CDED(10).DAmt; dlm; Check(1).CDED(10).YTDDAmt; dlm;
      

      '                         65                      66
      Print #CPFNHandle, Check(1).NetPay; dlm; Check(1).YTDNetPay; dlm;
      
      '                                 67                               68                               69
      Print #CPFNHandle, QPTrim$(Check(1).CDED(11).DCode); dlm; Check(1).CDED(11).DAmt; dlm; Check(1).CDED(11).YTDDAmt; dlm;
      
      
      '                           70                        71
      Print #CPFNHandle, TransRec(1).StaTaxAmt; dlm; Check(1).YTDState; dlm;
      
      
      '                                72                                  73                           74
      Print #CPFNHandle, QPTrim$(Check(1).CDED(12).DCode); dlm; Check(1).CDED(12).DAmt; dlm; Check(1).CDED(12).YTDDAmt; dlm;
      
      
      '                             75                                         76
      Print #CPFNHandle, Check(1).TotAdditEarn; dlm; OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn); dlm;
      
      '                         77                    78                                       79
      Print #CPFNHandle, Check(1).DDFlag; dlm; Check(1).CheckNum; dlm; SpellNumber$(Using$("####0.00", Check(1).NetPay)); dlm;
      '                           80                          81                         82                              83
      Print #CPFNHandle, MakeRegDate(CheckDate); dlm; Check(1).NetPay; dlm; QPTrim$(Check(1).EmpName); dlm; QPTrim$(Check(1).EmpAddr1); dlm;
      
      CityStateZip$ = QPTrim$(Check(1).EmpCity) + "  " + QPTrim$(Check(1).EmpState) + "  " + QPTrim$(Check(1).EmpZip)
      '                       84                            85                               86                         87                                        88                                     89
      Print #CPFNHandle, CityStateZip$; dlm; Str(Check(1).HolUsed + Check(1).PerUsed); dlm; Str(HolPay# + PerPay#); dlm; Str(HAmt# + PAmt#); dlm; Str(Check(1).HolUsed + Check(1).PerUsed); dlm; Str(HOLBAL# + PERBAL#)
  Return
  
LaserTopPrint:
      RPTSetupPRN 15, CPFNHandle '7/20
      If Check(1).DDFlag = True Then
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(67); Using("######", Check(1).CheckNum)
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(6); SpellNumber$(Using$("####0.00", Check(1).NetPay))  'Print the whole number part
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName); Tab(48); MakeRegDate(CheckDate); Tab(63); Using("$#,###,0.00", Check(1).NetPay)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, Tab(10); "VOID            VOID             VOID            VOID"
        Print #CPFNHandle, ""
        Print #CPFNHandle, "~"
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
      Else
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(67); Using("######", Check(1).CheckNum)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(6); SpellNumber$(Using$("####0.00", Check(1).NetPay))  'Print the whole number part
        Print #CPFNHandle,
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpName); Tab(48); MakeRegDate(CheckDate); Tab(63); Using("$#,###,0.00", Check(1).NetPay)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpAddr1)
        Print #CPFNHandle, Tab(12); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle,
        Print #CPFNHandle, ""
        Print #CPFNHandle, "~"
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
        Print #CPFNHandle, ""
      End If
      Print #CPFNHandle, "~ "; QPTrim$(Check(1).EmpName)
      Print #CPFNHandle, "   EMP NO.         SSN        PERIOD END    CHK DATE    BASE RATE"
      Print #CPFNHandle, QPTrim$(Check(1).EmpNo);
      Print #CPFNHandle, Tab(16); Check(1).EmpSSN;
      Print #CPFNHandle, Tab(31); MakeRegDate(Check(1).PayEndDate); Tab(43); MakeRegDate(CheckDate);
      Print #CPFNHandle, Tab(56); Using("####0.00", Check(1).BaseRate)
      Print #CPFNHandle, "        HRS    Period       YTD   Deductions    Period     YTD  BALANCE"
      Print #CPFNHandle, "REG";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).RegHrsWork);
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound(Check(1).RegHrsWork * Check(1).BaseRate));  'TAB(25); Check(1).YTDGrossPay;
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(1).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(1).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(1).YTDDAmt);
      Print #CPFNHandle, "  VAC  "; Using("###0.00", Check(1).VactBal)

      Print #CPFNHandle, "OT";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).OTHrsPaid);
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotOTWage);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(2).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(2).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(2).YTDDAmt);
      Print #CPFNHandle, "  SICK "; Using("###0.00", Check(1).SickBal)

      Print #CPFNHandle, "VAC";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", VacPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(3).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(3).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(3).YTDDAmt);
      Print #CPFNHandle, "  COMP "; Using("###0.00", Check(1).CompBal)

      Print #CPFNHandle, "SICK";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", SickPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(4).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(4).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(4).YTDDAmt)

      Print #CPFNHandle, "HOL";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).HolUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", HolPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(5).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(5).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(5).YTDDAmt)

      Print #CPFNHandle, "COMP";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", CompPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(6).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(6).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(6).YTDDAmt)

      Print #CPFNHandle, "GROSS";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).GrossPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDGrossPay);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(7).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(7).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(7).YTDDAmt)

      Print #CPFNHandle, "FWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).FedTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDFederal);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(8).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(8).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(8).YTDDAmt)

      Print #CPFNHandle, "FICA";
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(9).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(9).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(9).YTDDAmt)

      Print #CPFNHandle, "RET";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).RetireAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDRetire);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(10).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(10).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(10).YTDDAmt)

      Print #CPFNHandle, "NET PAY";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).NetPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDNetPay);
      Print #CPFNHandle, Tab(35); Check(1).CDED(11).DCode;
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(11).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(11).YTDDAmt)
      Print #CPFNHandle, "SWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", TransRec(1).StaTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDState);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(12).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(12).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(12).YTDDAmt)
      Print #CPFNHandle, "TOT ADD";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotAdditEarn);
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn))
      Print #CPFNHandle, ""
      Print #CPFNHandle, ""
      Print #CPFNHandle, ""
      Print #CPFNHandle, QPTrim$(Check(1).EmpName)
      Print #CPFNHandle, "   EMP NO.         SSN        PERIOD END    CHK DATE    BASE RATE"
      Print #CPFNHandle, QPTrim$(Check(1).EmpNo);
      Print #CPFNHandle, Tab(16); QPTrim$(Check(1).EmpSSN);
      Print #CPFNHandle, Tab(31); MakeRegDate(Check(1).PayEndDate); Tab(43); MakeRegDate(CheckDate);
      Print #CPFNHandle, Tab(56); Using("####0.00", Check(1).BaseRate)
      Print #CPFNHandle, "        HRS    Period        YTD  Deductions    Period     YTD  BALANCE"
      Print #CPFNHandle, "REG";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).RegHrsWork);
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound(Check(1).RegHrsWork * Check(1).BaseRate));  ' TAB(25); Check(1).YTDGrossPay;
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(1).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(1).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(1).YTDDAmt);
      Print #CPFNHandle, "  VAC  "; Using("###0.00", Check(1).VactBal)

      Print #CPFNHandle, "OT";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).OTHrsPaid);
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotOTWage);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(2).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(2).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(2).YTDDAmt);
      Print #CPFNHandle, "  SICK "; Using("###0.00", Check(1).SickBal)

      Print #CPFNHandle, "VAC";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).VacUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", VacPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(3).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(3).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(3).YTDDAmt);
      Print #CPFNHandle, "  COMP "; Using("###0.00", Check(1).CompBal)

      Print #CPFNHandle, "SICK";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).SickUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", SickPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(4).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(4).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(4).YTDDAmt)

      Print #CPFNHandle, "HOL";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).HolUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", HolPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(5).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(5).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(5).YTDDAmt)

      Print #CPFNHandle, "COMP";
      Print #CPFNHandle, Tab(6); Using("##0.00", Check(1).CompUsed);
      Print #CPFNHandle, Tab(14); Using("####0.00", CompPay#);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(6).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(6).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(6).YTDDAmt)

      Print #CPFNHandle, "GROSS";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).GrossPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDGrossPay);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(7).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(7).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(7).YTDDAmt)

      Print #CPFNHandle, "FWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).FedTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDFederal);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(8).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(8).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(8).YTDDAmt)

      Print #CPFNHandle, "FICA";
      Print #CPFNHandle, Tab(14); Using("####0.00", OldRound#(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
      Print #CPFNHandle, Tab(25); Using("####0.00", OldRound#(Check(1).YTDSocial + Check(1).YTDMedicare));
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(9).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(9).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(9).YTDDAmt)

      Print #CPFNHandle, "RET";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).RetireAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDRetire);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(10).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(10).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(10).YTDDAmt)

      Print #CPFNHandle, "NET PAY";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).NetPay);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDNetPay);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(11).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(11).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(11).YTDDAmt)
      Print #CPFNHandle, "SWT";
      Print #CPFNHandle, Tab(14); Using("####0.00", TransRec(1).StaTaxAmt);
      Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).YTDState);
      Print #CPFNHandle, Tab(35); QPTrim$(Check(1).CDED(12).DCode);
      Print #CPFNHandle, Tab(48); Using("###0.00", Check(1).CDED(12).DAmt);
      Print #CPFNHandle, Tab(56); Using("###0.00", Check(1).CDED(12).YTDDAmt)
      Print #CPFNHandle, "TOT ADD";
      Print #CPFNHandle, Tab(14); Using("####0.00", Check(1).TotAdditEarn);
      Print #CPFNHandle, "~"; Tab(24); Using("####0.00", OldRound(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn)); Chr$(12); '7/24 added ~ and adjusted Tab to allow for it
'      Print #CPFNHandle, "~"; Chr$(12);


  Return
ErrorHandler:
  Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmSPRTThisEmp", "CheckForValidWHNum", Erl)
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
'  Close
'  MsgBox "A folder named PRRPTS may be missing. Call Southern Software at 1-800-842-8190"
  
End Sub

Sub CreateCheckRegisterG()
  Dim Title$
  Dim NumOfChecks As Long
  Dim TotChecksAmt#
  Dim IdxRecLen As Integer
  Dim CRecLen As Integer
  Dim NumOfRecs As Long
  Dim Unit(1) As UnitFileRecType
  Dim UHandle As Integer
  Dim Image1$, Image2$
  Dim NextChkNum&
  Dim RHandle As Integer
  Dim CHandle As Integer
  Dim cnt As Long, GRecNum&
  Dim CheckDate As Integer
  Dim ToPrint$, x As Integer
  Dim IdxNHandle As Integer
  Dim IdxNRec As NumbSortIdxType
  Dim UTemp$, RptTitle$
  Dim CheckRegisterRptName$
  Dim NumOfDrafts As Integer
  Dim AmtOfDrafts As Double
  Dim AmtOfChecks As Double
  Dim dlm$, Draft$, ARName$
    
  dlm$ = "~"
  Title$ = "Updating Check Register"
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show ' , Me
  NumOfDrafts = 0
  NumOfChecks = 0
  TotChecksAmt# = 0

  ReDim Check(1) As PRCheckRecType

  IdxRecLen = 2
  CRecLen = Len(Check(1))
  
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) / IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle

  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  
  ReDim DrftNum(1) As String * 10
  ReDim ChkNum(1) As String * 10
  ReDim ChkNet(1) As String * 14

  Image1$ = "$###,##0.00"
  Image2$ = "######"

  NextChkNum& = -1
  
  RHandle = FreeFile
  ARName = "prrpts\Errorchk.RPT"
  On Error GoTo ErrorHandler
  Open ARName For Output As RHandle 'making sure the PRRPTS folder
  'exists because if it doesn't the program will crash
  Close RHandle
  
  RHandle = FreeFile
  CheckRegisterRptName = "prrpts\CHECKREGG.RPT"
  KillFile "prrpts\CHECKREGG.RPT"
  Open CheckRegisterRptName For Output As RHandle
  OpenChecksFile CHandle
  cnt = 0
  Do
    cnt = cnt + 1
    Get CHandle, cnt, Check(1)
  Loop Until Check(1).CActive = True Or (cnt > NumOfRecs)

  CheckDate = Check(1).CheckDate

  For cnt = 1 To NumOfRecs
    GRecNum& = CLng(IdxBuff(cnt))
    Get CHandle, GRecNum&, Check(1)
    If Check(1).CActive = True Then
      If NextChkNum& = -1 Then
        NextChkNum& = Check(1).CheckNum
      Else
        NextChkNum& = NextChkNum& + 1
      End If
DoThisOne:
      If NextChkNum& = Check(1).CheckNum Then
        If Check(1).DDFlag = True Then 'added 9/10
          NumOfDrafts = NumOfDrafts + 1
          AmtOfDrafts = OldRound#(AmtOfDrafts + Check(1).NetPay)
        Else
          NumOfChecks = NumOfChecks + 1
          AmtOfChecks = OldRound#(AmtOfChecks + Check(1).NetPay)
        End If
        TotChecksAmt# = OldRound#(TotChecksAmt# + Check(1).NetPay)
        Draft$ = ""
        If Check(1).DDFlag = True Then
          Draft$ = "Draft"
        Else
          Draft$ = ""
        End If
        '                      0                        1
        Print #RHandle, Unit(1).UFEMPR; dlm; MakeRegDate(CheckDate); dlm;
        '                       2                   3                    4
        Print #RHandle, Check(1).CheckNum; dlm; Check(1).EmpName; dlm; Draft; dlm;
        '                              5
        Print #RHandle, Using$(Image1$, Check(1).NetPay); dlm;
        '                6            7                 8                 9
        Print #RHandle, ""; dlm; NumOfChecks; dlm; AmtOfChecks; dlm; NumOfDrafts; dlm;
        '                    10                       11                             12
        Print #RHandle, AmtOfDrafts; dlm; AmtOfChecks + AmtOfDrafts; dlm; NumOfDrafts + NumOfChecks
      Else
        '                      0                       1                   2
        Print #RHandle, Unit(1).UFEMPR; dlm; MakeRegDate(CheckDate); dlm; ""; dlm;
        '               3        4        5                 6                     7
        Print #RHandle, ""; dlm; ""; dlm; ""; dlm; "****REPRINTED****"; dlm; NumOfChecks; dlm;
        '                   8                  9                 10                       11
        Print #RHandle, AmtOfChecks; dlm; NumOfDrafts; dlm; AmtOfDrafts; dlm; AmtOfChecks + AmtOfDrafts; dlm;
        '                         12
        Print #RHandle, NumOfDrafts + NumOfChecks

        NextChkNum& = Check(1).CheckNum
        GoTo DoThisOne
      End If
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close CHandle
  Close RHandle
  RptTitle$ = "Payroll Check Register Report"
  arCheckRegister.Show
  frmLoadingRpt.Show
  MainLog ("Check register printed.")
Exit Sub

ErrorHandler:
  Close
  Unload FrmShowPctComp
  MsgBox "ERROR: A folder named PRRPTS may be needed."

End Sub

Public Sub UnloadAllFormsAndOpn(RegExit As Boolean)
  
  Unload frmChkPrintInfo
  Unload frmChkPrintingMenu
  Unload frmChkReprintInfo
  Unload frmLoadingRpt
  Unload frmPrint
  Unload frmPrintChks
  Unload FrmShowPctComp
  Unload frmViewPrintChks
  Unload frmWarnBadCheckNum
  Unload frmWarnFilesMissed
  Unload arCheckRegister
  If PWcnt = -3 Then Exit Sub
  If RegExit = False Then
    Call ClearInUse(PWcnt)
  End If
  
End Sub

Sub RPTSetupPRN(RPTNum, Handle)
  Dim RPTPitch As Integer
  Dim PrinterSetUpFile As Integer
  Dim PrntType As PRNSetupRecType
  Dim x As Integer
  Dim PHandle As Integer
  Dim DefPrinter As String
  Dim PrnDef As String
  Dim LineLen As Integer
  Dim TextLine$
  Dim y As Integer
  Dim z As Integer
  Dim NextCommaPOS As Integer
  Dim CodeStartPOS As Integer
  Dim Codeline1$
  Dim Codeline2$
  'this sub coordinates the printing procedure so that any
  'pitch data saved in the Printer setup screen for a
  'particular report gets sent to the printer
  For z = 1 To 10
    ToPrint1(z) = 0
    ToPrint2(z) = 0
  Next z
  OpenPrinterSetupFile PrinterSetUpFile
  Get PrinterSetUpFile, 1, PrntType
  Close PrinterSetUpFile
  DefPrinter = QPTrim$(PrntType.Printer)
  'if a pitch isn't saved for this print job then by
  'default the pitch becomes 10
  
  If RPTNum = 123 Then GoTo SkipThis
  
  RPTPitch = PrntType.RPT(RPTNum)
  
SkipThis:

  GoSub GetPrinterCodes
  If Len(Codeline1) Then
    Select Case y
      Case 1:
        Print #Handle, Chr(ToPrint1(1));
      Case 2:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2));
      Case 3:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3));
      Case 4:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4));
      Case 5:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5));
      Case 6:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6));
      Case 7:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7));
      Case 8:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8));
      Case 9:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9));
      Case 10:
        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9)); Chr(ToPrint1(10));
      Case Else:
    End Select
  ElseIf Len(Codeline2) Then
    Select Case y
      Case 1:
        Print #Handle, Chr(ToPrint2(1));
      Case 2:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2));
      Case 3:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3));
      Case 4:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4));
      Case 5:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5));
      Case 6:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6));
      Case 7:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7));
      Case 8:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8));
      Case 9:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9));
      Case 10:
        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9)); Chr(ToPrint2(10));
      Case Else:
    End Select
  End If
  
  Exit Sub
  
GetPrinterCodes:
  PHandle = FreeFile
  Open "PRData\Prprndf.dat" For Input As #PHandle  ' Open file.
  Line Input #PHandle, TextLine   ' Read first line into TextLine.
   'the second line is where individual printers start their codes
   NextCommaPOS = 1
   
   Do While Not eof(PHandle) And NextCommaPOS <> 0  ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
     If TextLine = "@" + DefPrinter$ Then 'if we locate the default printer
     'printer
         If eof(PHandle) Then Exit Do 'if for some reason we get to the end of the file
         'then exit
         If RPTNum = 123 Then
           Line Input #PHandle, TextLine 'read next line
             LineLen = Len(TextLine)
             Codeline1 = Mid(TextLine, 11, LineLen)
             CodeStartPOS = 1
             y = 1
             Do
               NextCommaPOS = InStr(CodeStartPOS, Codeline1, ",")
               If NextCommaPOS = 0 Then
                 LineLen = Len(Codeline1)
                 ToPrint1(y) = CInt(Mid(Codeline1, CodeStartPOS, 3))
                 Exit Do
               End If
               ToPrint1(y) = CInt(Mid(Codeline1, CodeStartPOS, 3))
               CodeStartPOS = NextCommaPOS + 1
               y = y + 1
             Loop Until NextCommaPOS = 0
             GoTo XIsOne
         End If
       Do
         Line Input #PHandle, TextLine
         If Mid(TextLine, 1, 2) = RPTPitch Then
           LineLen = Len(TextLine)
           Codeline2 = Mid(TextLine, 11, LineLen)
           y = 1
           CodeStartPOS = 1
           Do
             NextCommaPOS = InStr(CodeStartPOS, Codeline2, ",")
             If NextCommaPOS = 0 Then
               LineLen = Len(Codeline2)
               ToPrint2(y) = CInt(Mid(Codeline2, CodeStartPOS, 3))
               Exit Do
             End If
             ToPrint2(y) = CInt(Mid(Codeline2, CodeStartPOS, 3))
             CodeStartPOS = NextCommaPOS + 1
             y = y + 1
           Loop Until NextCommaPOS = 0
           
           Exit Do
         End If
        Loop Until NextCommaPOS = 0
XIsOne:
     End If 'ends if TextLine = @ + DefPrinter
   Loop
   Close #PHandle   ' Close file.
   Return
   
End Sub


Public Function FilesROK(frm As Form, InFileNames() As String, OutFileNames() As String, ThisMany As Integer) As Boolean
  Dim NextName As Integer
  Dim x As Integer
  FilesROK = True
  NextName = 1
  For x = 1 To ThisMany
    If Not Exist(InFileNames(x)) Then
      OutFileNames(NextName) = InFileNames(x)
      NextName = NextName + 1
      FilesROK = False
    End If
  Next x
  If FilesROK = False Then
    frmWarnFilesMissed.Show vbModal, frm
    For x = 1 To ThisMany
      InFileNames(x) = ""
      OutFileNames(x) = ""
    Next x
  End If
End Function

Private Sub GetBenefits(SAmt As Double, VAmt As Double, HAmt As Double, PAmt As Double, EmpRec2 As EmpData2Type)

  Dim HireDate As Integer
  Dim WhatLeaveTbl As Integer
  Dim AccrualDays As Integer
  Dim YearsOfService As Integer
  Dim cnt As Integer
  Dim VTableEntry As Integer
  Dim HTableEntry As Integer
  Dim PTableEntry As Integer
  
'  Dim VADJFlag As Boolean
  Dim StableEntry As Integer
'  Dim SADJFlag As Boolean
  Dim LastDate As Integer, x As Integer
  Dim AccrueHandle As Long
  Dim AccrualDate As Integer
  Dim LeaveHandle As Integer
  Dim NumLeaveRec As Integer
  ReDim LeaveRec(1) As LeaveRecType

  OpenLeaveFileName LeaveHandle
  NumLeaveRec = LOF(LeaveHandle) \ Len(LeaveRec(1))
'  If NumLeaveRec = 0 Then
'    Unload FrmShowPctComp
'    MsgBox "No records on file"
'    Close
'    Exit Sub
'  End If
  ReDim LeaveRec(1 To NumLeaveRec) As LeaveRecType

  For x = 1 To NumLeaveRec
    Get LeaveHandle, x, LeaveRec(x)
  Next x
  Close LeaveHandle

  AccrueHandle = FreeFile
  Open PRData + AccrueFileName For Random Shared As AccrueHandle Len = 2
  Get AccrueHandle, 1, LastDate
  Close AccrueHandle
'  AccrualDate = ReplaceString(MakeRegDate(LastDate), "/", "-")
  AccrualDate = LastDate
  HireDate = EmpRec2.EMPHDATE
'  If HireDate <= -11000 Then 'roughly 1950
'    GoSub BadHireDate
'    GoTo BadDateSkip
'  End If

  WhatLeaveTbl = EmpRec2.LeaveTbl 'get data from
  'leave table assigned to this employee
  If WhatLeaveTbl < 1 Then
    WhatLeaveTbl = 1
  End If
  AccrualDays = AccrualDate - HireDate
  If AccrualDays > 365 Then
    YearsOfService = Int(AccrualDays / 365)
  Else
    YearsOfService = 0
  End If
  
  For cnt = 1 To 20
    If YearsOfService <= LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
      Exit For
    End If
  Next
  If cnt > 20 Then cnt = 20
  If YearsOfService = LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
    VTableEntry = cnt
  Else
    VTableEntry = cnt - 1
  End If
  If VTableEntry = 0 Then VTableEntry = 1
  VAmt# = OldRound#(LeaveRec(WhatLeaveTbl).VEntry(VTableEntry).EARN * (EmpRec2.EMPBCODE * 0.01))
  If VAmt# > 0 Then           ' if there is amount to add
    If EmpRec2.EMPVBAL + VAmt# > LeaveRec(WhatLeaveTbl).VacMax Then     'if > max amt
      VAmt# = LeaveRec(WhatLeaveTbl).VacMax - EmpRec2.EMPVBAL   'set amt to max
    End If                                             '
  End If
  
  For cnt = 1 To 20
    If YearsOfService <= LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
      Exit For
    End If
  Next
  If cnt > 20 Then cnt = 20
  If YearsOfService = LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
    StableEntry = cnt
  Else
    StableEntry = cnt - 1
  End If
  If StableEntry = 0 Then StableEntry = 1
  SAmt# = OldRound#(LeaveRec(WhatLeaveTbl).SEntry(StableEntry).EARN * (EmpRec2.EMPBCODE * 0.01)) '8/5
  If SAmt# > 0 Then           ' if there is amount to add
    If EmpRec2.EMPSLBAL + SAmt# > LeaveRec(WhatLeaveTbl).SICKMAX Then   'if > max amt
      SAmt# = LeaveRec(WhatLeaveTbl).SICKMAX - EmpRec2.EMPSLBAL
    End If
  End If

  For cnt = 1 To 20
    If YearsOfService <= LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
      Exit For
    End If
  Next
  If cnt > 20 Then cnt = 20
  If YearsOfService = LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
    HTableEntry = cnt
  Else
    HTableEntry = cnt - 1
  End If
  If HTableEntry = 0 Then HTableEntry = 1
  HAmt# = OldRound#(LeaveRec(WhatLeaveTbl).HEntry(HTableEntry).EARN * (EmpRec2.EMPBCODE * 0.01)) '8/5
  If HAmt# > 0 Then           ' if there is amount to add
    If EmpRec2.HOLBAL + HAmt# > LeaveRec(WhatLeaveTbl).HolMax Then     'if > max amt
      HAmt# = LeaveRec(WhatLeaveTbl).HolMax - EmpRec2.HOLBAL
    End If
  End If

  For cnt = 1 To 20
    If YearsOfService <= LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
      Exit For
    End If
  Next
  If cnt > 20 Then cnt = 20
  If YearsOfService = LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
    PTableEntry = cnt
  Else
    PTableEntry = cnt - 1
  End If
  If PTableEntry = 0 Then PTableEntry = 1
  PAmt# = OldRound#(LeaveRec(WhatLeaveTbl).PEntry(PTableEntry).EARN * (EmpRec2.EMPBCODE * 0.01)) '8/5
  If PAmt# > 0 Then           ' if there is amount to add
    If EmpRec2.PERBAL + PAmt# > LeaveRec(WhatLeaveTbl).PerMax Then     'if > max amt
      PAmt# = LeaveRec(WhatLeaveTbl).PerMax - EmpRec2.PERBAL
    End If
  End If

End Sub

Sub CreateCheckRegisterT()
  Dim Title$
  Dim NumOfChecks As Long
  Dim TotChecksAmt#
  Dim IdxRecLen As Integer
  Dim CRecLen As Integer
  Dim NumOfRecs As Long
  Dim Unit(1) As UnitFileRecType
  Dim UHandle As Integer
  Dim Image1$, Image2$
  Dim LineCnt As Integer
  Dim NextChkNum&, MaxLines As Integer
  Dim Page As Integer
  Dim RHandle As Integer
  Dim CHandle As Integer
  Dim cnt As Long, GRecNum&
  Dim CheckDate As Integer
  Dim ToPrint$, x As Integer
  Dim IdxNHandle As Integer
  Dim IdxNRec As NumbSortIdxType
  Dim Dash As String * 70
  Dim UTemp$, RptTitle$, FF$
  Dim CheckRegisterRptName$
  Dim NumOfDrafts As Integer
  Dim AmtOfDrafts As Double
  Dim AmtOfChecks As Double
  
  Title$ = "Updating Check Register"
  FrmShowPctComp.Label1 = Title$
  FrmShowPctComp.Show ' , Me
  FF$ = Chr$(12)
  NumOfDrafts = 0
  NumOfChecks = 0
  TotChecksAmt# = 0

  ReDim Check(1) As PRCheckRecType
  ReDim Pg(1) As String * 3

  IdxRecLen = 2
  CRecLen = Len(Check(1))
  
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) / IdxRecLen
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle

  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  
  ReDim DrftNum(1) As String * 10
  ReDim ChkNum(1) As String * 10
  ReDim ChkNet(1) As String * 14

  Image1$ = "$###,##0.00"
  Image2$ = "######"

  Dash = String$(70, "-")

  LineCnt = 0
  NextChkNum& = -1
  MaxLines = 55
  Page = 0
  
  RHandle = FreeFile
  CheckRegisterRptName = "prrpts\CHECKREG.RPT"
  KillFile "prrpts\CHECKREG.RPT"
  Open CheckRegisterRptName For Output As RHandle
  RPTSetupPRN 16, RHandle
  OpenChecksFile CHandle
  cnt = 0
  Do
    cnt = cnt + 1
    Get CHandle, cnt, Check(1)
'    If DosError Then Exit Do
  Loop Until Check(1).CActive = True Or (cnt > NumOfRecs)

  CheckDate = Check(1).CheckDate

  GoSub PrintCheckHeader

  For cnt = 1 To NumOfRecs
    GRecNum& = CLng(IdxBuff(cnt))
    Get CHandle, GRecNum&, Check(1)

    If Check(1).CActive = True Then
      ToPrint$ = Space$(70)
      If NextChkNum& = -1 Then
        NextChkNum& = Check(1).CheckNum
      Else
        NextChkNum& = NextChkNum& + 1
      End If
DoThisOne:
      If NextChkNum& = Check(1).CheckNum Then
        If Check(1).DDFlag = True Then 'added 9/10
          NumOfDrafts = NumOfDrafts + 1
          AmtOfDrafts = OldRound#(AmtOfDrafts + Check(1).NetPay)
        Else
          NumOfChecks = NumOfChecks + 1
          AmtOfChecks = OldRound#(AmtOfChecks + Check(1).NetPay)
        End If
        TotChecksAmt# = OldRound#(TotChecksAmt# + Check(1).NetPay)
        RSet ChkNum(1) = Using$(Image2$, Check(1).CheckNum)
        LSet ToPrint$ = ChkNum(1)
        If Check(1).DDFlag = True Then
          Mid$(ToPrint$, 15) = Check(1).EmpName + "Draft"
        Else
          Mid$(ToPrint$, 15) = Check(1).EmpName
        End If
        Mid$(ToPrint$, 59) = Using$(Image1$, Check(1).NetPay)
        Print #RHandle, ToPrint$
        LineCnt = LineCnt + 1
        If LineCnt >= MaxLines Then
          Print #RHandle, FF$
          GoSub PrintCheckHeader
          
        End If
      Else
        LSet ToPrint$ = "              ****REPRINTED****"
        Print #RHandle, ToPrint$
        LineCnt = LineCnt + 1
        NextChkNum& = Check(1).CheckNum
        GoTo DoThisOne
      End If
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  GoSub PrintTotChecksLine

  Close CHandle
  RPTSetupPRN 123, RHandle
  Close RHandle
  UTemp$ = "": ToPrint$ = ""
  RptTitle$ = "Payroll Check Register Report"
  ViewPrint CheckRegisterRptName, RptTitle$, True, 1, False, 1
  MainLog ("Check register printed.")
Exit Sub

PrintCheckHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(70)
  LSet UTemp$ = Unit(1).UFEMPR
  Mid$(UTemp$, 62) = "Page:" + Pg(1)
  Print #RHandle, UTemp$
  Print #RHandle, "Payroll Check Register"
  Print #RHandle, "Check Date: " + MakeRegDate(CheckDate)
  Print #RHandle,
  Print #RHandle, " Check No.    Employee Name                               Check Amount"
  Print #RHandle, Dash
  LineCnt = 6
Return

PrintTotChecksLine:
  LSet ChkNet(1) = LTrim$(Using(Image1$, TotChecksAmt#))
  LSet ChkNum(1) = LTrim$(Using(Image2$, NumOfChecks))
  If NumOfDrafts > 0 Then
    LSet DrftNum(1) = LTrim$(Using(Image2$, NumOfDrafts))
  End If
  Print #RHandle,
  Print #RHandle, Dash
  Print #RHandle,
  Print #RHandle, " Summary:"
  Print #RHandle,
  Print #RHandle, "    Number of Checks Printed: " + ChkNum(1)
  If NumOfDrafts > 0 Then
    Print #RHandle, "    Total Amount of Checks: " + Using(Image1$, AmtOfChecks)
    Print #RHandle,
    Print #RHandle, "    Number of Drafts Printed: " + DrftNum(1)
    Print #RHandle, "    Total Amount of Drafts: " + Using(Image1, AmtOfDrafts)
    Print #RHandle,
    Print #RHandle, "    Total Paid: " & (NumOfDrafts + NumOfChecks)
  End If
  If NumOfDrafts > 0 Then
    Print #RHandle, "    Total Amount of Checks and Drafts: " + ChkNet(1)
  Else
    Print #RHandle, "    Total Amount of Checks: " + ChkNet(1)
  End If
  Print #RHandle,
  Print #RHandle, FF$
Return
End Sub

Public Sub GetVoidChkData()
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim NumOfVoids As Integer
  Dim ChkRec As PRCheckRecType
  Dim ChkHandle As Integer
  Dim x As Integer, y As Integer, z As Integer
  Dim NumOfChks As Integer
  Dim ChkCnt As Integer
  Dim VoidCnt As Integer
  Dim ThisEmpNum$
  Dim ThisCnt As Integer
  Dim EVdCnt As Integer
  
  OpenChecksFile ChkHandle
  NumOfChks = LOF(ChkHandle) / Len(ChkRec)
  Dim NumAccts As Integer

  OpenTempVoidFile TVHandle
  NumOfVoids = LOF(TVHandle) / Len(TempVoid)
'  NumAccts = LOF(TVHandle) / Len(TempVoid)
'
'  For x = 1 To NumAccts
'    Get TVHandle, x, TempVoid
'      Debug.Print TempVoid.PRNetGL + " PRNET            " + CStr(TempVoid.PRNet)
'      Debug.Print TempVoid.SOCWHGL + " SOC Withholdings " + CStr(TempVoid.SOCWHAmt)
'      Debug.Print TempVoid.MEDWHGL + " MED Withholdings " + CStr(TempVoid.MEDWHAmt)
'      Debug.Print TempVoid.SOCMATCRGL + " SOC Match Liab   " + CStr(TempVoid.SOCMATCRAmt)
'      Debug.Print TempVoid.MEDMATCRGL + " MED Match Liab   " + CStr(TempVoid.MEDMATCRAmt)
'      Debug.Print TempVoid.FEDWHGL + " FED Withholdings " + CStr(TempVoid.FEDWHAmt)
'      Debug.Print TempVoid.STAWHGL + " STA Withholdings " + CStr(TempVoid.STAWHAmt)
'      Debug.Print TempVoid.RETWHGL + " RET Withholdings " + CStr(TempVoid.RETWHAmt)
'      Debug.Print TempVoid.RETMATCRGL + " RET Match Liab   " + CStr(TempVoid.RETMATCRAmt)
'      For y = 1 To 50
'        If TempVoid.DedData(y).DAmt > 0 Then
'          Debug.Print TempVoid.DedData(y).DedGLNum + " Deduction        " + CStr(TempVoid.DedData(y).DAmt)
'        End If
'      Next y
'      Debug.Print TempVoid.WagesGL + "  Wages           " + CStr(TempVoid.WagesAmt)
'      Debug.Print TempVoid.SOCMATDBGL + " SOC Match        " + CStr(TempVoid.SOCMATDBAmt)
'      Debug.Print TempVoid.MEDMATDBGL + " MED Match        " + CStr(TempVoid.MEDMATDBAmt)
'      Debug.Print TempVoid.RETMATDBGL + " RET Match        " + CStr(TempVoid.RETMATDBAmt)
'      Debug.Print TempVoid.CheckAmt
'      Debug.Print TempVoid.CheckDate
'      Debug.Print TempVoid.CheckNum
'  Next x
'  Close TVHandle
  
  
  For x = 1 To NumOfChks
    Get ChkHandle, x, ChkRec
      If ChkRec.CActive = True Then
        ThisEmpNum = QPTrim$(ChkRec.EmpNo)
        If ThisEmpNum = "" Then GoTo NextEmp
        EVdCnt = 0
        ReDim GetThese(1 To 1) As Long
        For y = 1 To NumOfVoids
          Get TVHandle, y, TempVoid
            If QPTrim$(TempVoid.EmpNum) = ThisEmpNum Then 'GoTo NextOne
              EVdCnt = EVdCnt + 1
              ReDim Preserve GetThese(1 To EVdCnt) As Long
              GetThese(EVdCnt) = y
            End If
        Next y
'            VoidCnt = TempVoid.NumOfAccts
'            ThisCnt = (y - 1) + VoidCnt
'            For z = y To ThisCnt
            For z = 1 To EVdCnt
'              Get TVHandle, z, TempVoid
              Get TVHandle, GetThese(z), TempVoid
                TempVoid.CheckAmt = ChkRec.NetPay
                TempVoid.CheckDate = ChkRec.CheckDate
                TempVoid.CheckNum = ChkRec.CheckNum
              Put TVHandle, GetThese(z), TempVoid
            Next z
'            GoTo NextEmp
NextOne:
'        Next y
    End If
NextEmp:
  Next x
  
  Close TVHandle
  Close ChkHandle
  

'  OpenTempVoidFile TVHandle
'  NumAccts = LOF(TVHandle) / Len(TempVoid)
'  For x = 1 To NumAccts
'    Get TVHandle, x, TempVoid
'      Debug.Print TempVoid.PRNetGL + " PRNET            " + CStr(TempVoid.PRNet)
'      Debug.Print TempVoid.SOCWHGL + " SOC Withholdings " + CStr(TempVoid.SOCWHAmt)
'      Debug.Print TempVoid.MEDWHGL + " MED Withholdings " + CStr(TempVoid.MEDWHAmt)
'      Debug.Print TempVoid.SOCMATCRGL + " SOC Match Liab   " + CStr(TempVoid.SOCMATCRAmt)
'      Debug.Print TempVoid.MEDMATCRGL + " MED Match Liab   " + CStr(TempVoid.MEDMATCRAmt)
'      Debug.Print TempVoid.FEDWHGL + " FED Withholdings " + CStr(TempVoid.FEDWHAmt)
'      Debug.Print TempVoid.STAWHGL + " STA Withholdings " + CStr(TempVoid.STAWHAmt)
'      Debug.Print TempVoid.RETWHGL + " RET Withholdings " + CStr(TempVoid.RETWHAmt)
'      Debug.Print TempVoid.RETMATCRGL + " RET Match Liab   " + CStr(TempVoid.RETMATCRAmt)
'      For y = 1 To 50
'        If TempVoid.DedData(y).DAmt > 0 Then
'          Debug.Print TempVoid.DedData(y).DedGLNum + " Deduction        " + CStr(TempVoid.DedData(y).DAmt)
'        End If
'      Next y
'      Debug.Print TempVoid.WagesGL + "  Wages           " + CStr(TempVoid.WagesAmt)
'      Debug.Print TempVoid.SOCMATDBGL + " SOC Match        " + CStr(TempVoid.SOCMATDBAmt)
'      Debug.Print TempVoid.MEDMATDBGL + " MED Match        " + CStr(TempVoid.MEDMATDBAmt)
'      Debug.Print TempVoid.RETMATDBGL + " RET Match        " + CStr(TempVoid.RETMATDBAmt)
'      Debug.Print TempVoid.CheckAmt
'      Debug.Print TempVoid.CheckDate
'      Debug.Print TempVoid.CheckNum
'  Next x
'  Close TVHandle
  
End Sub
Public Sub Print42LineCommon(CPFNHandle As Integer, Check() As PRCheckRecType, WorkPay#, HolPay#, SickPay#, CompPay#, VacPay#, TCheckNum&, CheckDate As Long)
'42 line Custom' Hamlet
  RPTSetupPRN 15, CPFNHandle
  If Check(1).DDFlag = False Then
    '--Line 1
    Print #CPFNHandle, "~";
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle, QPTrim$(Check(1).EmpNo);
    Print #CPFNHandle, Tab(12); MakeRegDate(Check(1).PayEndDate);
    Print #CPFNHandle, Tab(25); Using("###0.00", Check(1).BaseRate)
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle, "                HRS         EARN           YTD                         PERIOD            YTD"
    '--Line 4
    Print #CPFNHandle, Tab(1); "HRS WORKED  "; Using("###0.00", Check(1).RegHrsWork);
    Print #CPFNHandle, Tab(25); Using("####0.00", WorkPay#);
    Print #CPFNHandle, Tab(51); "RETIREMENT";
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).RetireAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).YTDRetire);

    '--Line 5
    Print #CPFNHandle, Tab(1); "HOL HRS     "; Using("###0.00", Check(1).HolUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", HolPay#);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(1).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(1).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(1).YTDDAmt)

    '--Line 6
    Print #CPFNHandle, Tab(1); "SICK HRS    "; Using("###0.00", Check(1).SickUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", SickPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).SickBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(2).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(2).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(2).YTDDAmt)

    '--Line 7
    Print #CPFNHandle, Tab(1); "COMP HRS    "; Using("###0.00", Check(1).CompUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", CompPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).CompBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(3).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(3).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(3).YTDDAmt);

    '--Line 8
    Print #CPFNHandle, Tab(1); "VAC HRS    "; Using(" ###0.00", Check(1).VacUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", VacPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).VactBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(4).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(4).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(4).YTDDAmt);

    '--Line 9
    Print #CPFNHandle, Tab(1); "TOT REG HRS "; Using("###0.00", Check(1).RegHrsPaid);
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotRegWage);
    'Print #CPFNHandle, TAB(39); USING "####0.00"; Round(EmpRec3(1).YTDRegWage);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(5).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(5).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(5).YTDDAmt);

    '--Line 10
    Print #CPFNHandle, Tab(1); "OT  HRS     "; Using("###0.00", Check(1).OTHrsPaid);
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotOTWage);
    'Print #CPFNHandle, TAB(39); USING "####0.00"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(6).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(6).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(6).YTDDAmt);

    '--Line 11
    Print #CPFNHandle, Tab(1); Check(1).AEarn(1).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(1).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(7).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(7).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(7).YTDDAmt)

    '--Line 12
    Print #CPFNHandle, Tab(1); Check(1).AEarn(2).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(2).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(8).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(8).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(8).YTDDAmt)

    '--Line 13
    Print #CPFNHandle, Tab(1); Check(1).AEarn(3).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(3).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(9).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(9).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(9).YTDDAmt)

    '--Line 14
    Print #CPFNHandle, Tab(1); "GROSS PAY";
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).GrossPay);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDGrossPay);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(10).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(10).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(10).YTDDAmt)

   '--Line 15
    Print #CPFNHandle, "FED TAX"; Tab(25); Using("####0.00", Check(1).FedTaxAmt);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDFederal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(11).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(11).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(11).YTDDAmt)

    '--Line 16
    Print #CPFNHandle, "STA TAX"; Tab(25); Using("####0.00", Check(1).StaTaxAmt);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDState);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(12).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(12).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(12).YTDDAmt)

    '--Line 17
    Print #CPFNHandle, "FICA"; Tab(25); Using("####0.00", OldRound(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
    Print #CPFNHandle, Tab(39); Using("####0.00", OldRound(Check(1).YTDSocial + Check(1).YTDMedicare));
    Print #CPFNHandle, Tab(51); "NET PAY";
    Print #CPFNHandle, Tab(70); Using("####0.00", Check(1).NetPay);
    Print #CPFNHandle, Tab(85); Using("####0.00", Check(1).YTDNetPay)

    '--Line 18 - Last line of stub

    '--Line 19 - First line of check
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(64); Using("######0", TCheckNum&);
      Print #CPFNHandle, Tab(74); MakeRegDate(CheckDate); Tab(84); Using("$##,##0.00", Check(1).NetPay)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(10); SpellNumber$(Using$("####0.00", Str$(Check(1).NetPay))) '--Line 30
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpName)
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpAddr1)
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, "~" '--Line 39
  Else
    '--Line 1
    Print #CPFNHandle, "~";
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle, QPTrim$(Check(1).EmpNo);
    Print #CPFNHandle, Tab(12); MakeRegDate(Check(1).PayEndDate);
    Print #CPFNHandle, Tab(25); Using("###0.00", Check(1).BaseRate)
    'Print #CPFNHandle, TAB(75); USING "Check No: ######"; TCheckNum&
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle,
    Print #CPFNHandle, "                HRS         EARN           YTD                         PERIOD            YTD"
    '--Line 4
    Print #CPFNHandle, Tab(1); "HRS WORKED  "; Using("###0.00", Check(1).RegHrsWork);
    Print #CPFNHandle, Tab(25); Using("####0.00", WorkPay#);
    Print #CPFNHandle, Tab(51); "RETIREMENT";
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).RetireAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).YTDRetire);

    '--Line 5
    Print #CPFNHandle, Tab(1); "HOL HRS     "; Using("###0.00", Check(1).HolUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", HolPay#);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(1).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(1).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(1).YTDDAmt)

    '--Line 6
    Print #CPFNHandle, Tab(1); "SICK HRS    "; Using("###0.00", Check(1).SickUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", SickPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).SickBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(2).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(2).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(2).YTDDAmt)

    '--Line 7
    Print #CPFNHandle, Tab(1); "COMP HRS    "; Using("###0.00", Check(1).CompUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", CompPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).CompBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(3).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(3).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(3).YTDDAmt);

    '--Line 8
    Print #CPFNHandle, Tab(1); "VAC HRS    "; Using(" ###0.00", Check(1).VacUsed);
    Print #CPFNHandle, Tab(25); Using("####0.00", VacPay#);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).VactBal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(4).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(4).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(4).YTDDAmt);

    '--Line 9
    Print #CPFNHandle, Tab(1); "TOT REG HRS "; Using("###0.00", Check(1).RegHrsPaid);
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotRegWage);
    'Print #CPFNHandle, TAB(39); USING "####0.00"; Round(EmpRec3(1).YTDRegWage);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(5).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(5).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(5).YTDDAmt);

    '--Line 10
    Print #CPFNHandle, Tab(1); "OT  HRS     "; Using("###0.00", Check(1).OTHrsPaid);
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).TotOTWage);
    'Print #CPFNHandle, TAB(39); USING "####0.00"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(6).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(6).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(6).YTDDAmt);

    '--Line 11
    Print #CPFNHandle, Tab(1); Check(1).AEarn(1).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(1).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(7).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(7).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(7).YTDDAmt)

    '--Line 12
    Print #CPFNHandle, Tab(1); Check(1).AEarn(2).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(2).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(8).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(8).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(8).YTDDAmt)

    '--Line 13
    Print #CPFNHandle, Tab(1); Check(1).AEarn(3).DCode;
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).AEarn(3).DAmt);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(9).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(9).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(9).YTDDAmt)

    '--Line 14
    Print #CPFNHandle, Tab(1); "GROSS PAY";
    Print #CPFNHandle, Tab(25); Using("####0.00", Check(1).GrossPay);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDGrossPay);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(10).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(10).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(10).YTDDAmt)

   '--Line 15
    Print #CPFNHandle, "FED TAX"; Tab(25); Using("####0.00", Check(1).FedTaxAmt);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDFederal);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(11).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(11).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(11).YTDDAmt)

    '--Line 16
    Print #CPFNHandle, "STA TAX"; Tab(25); Using("####0.00", Check(1).StaTaxAmt);
    Print #CPFNHandle, Tab(39); Using("####0.00", Check(1).YTDState);
    Print #CPFNHandle, Tab(51); QPTrim$(Check(1).CDED(12).DCode);
    Print #CPFNHandle, Tab(71); Using("###0.00", Check(1).CDED(12).DAmt);
    Print #CPFNHandle, Tab(86); Using("###0.00", Check(1).CDED(12).YTDDAmt)

    '--Line 17
    Print #CPFNHandle, "FICA"; Tab(25); Using("####0.00", OldRound(Check(1).MedTaxAmt + Check(1).SocTaxAmt));
    Print #CPFNHandle, Tab(39); Using("####0.00", OldRound(Check(1).YTDSocial + Check(1).YTDMedicare));
    Print #CPFNHandle, Tab(51); "NET PAY";
    Print #CPFNHandle, Tab(70); Using("####0.00", Check(1).NetPay);
    Print #CPFNHandle, Tab(85); Using("####0.00", Check(1).YTDNetPay)

    '--Line 18 - Last line of stub

    '--Line 19 - First line of check
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, "VOID                VOID                VOID                  VOID"
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(64); Using("######0", TCheckNum&);
      Print #CPFNHandle, Tab(74); MakeRegDate(CheckDate); Tab(84); Using("$##,##0.00", Check(1).NetPay)
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(10); SpellNumber$(Using$("####0.00", Str$(Check(1).NetPay))) '--Line 30
      Print #CPFNHandle, "VOID                VOID                VOID                  VOID"
      Print #CPFNHandle,
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpName)
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpAddr1)
      Print #CPFNHandle, Tab(14); QPTrim$(Check(1).EmpCity); " "; QPTrim$(Check(1).EmpState); " "; QPTrim$(Check(1).EmpZip)
      Print #CPFNHandle,
      Print #CPFNHandle, "VOID                VOID                VOID                  VOID"
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle,
      Print #CPFNHandle, "~" '--Line 39
  End If

End Sub


