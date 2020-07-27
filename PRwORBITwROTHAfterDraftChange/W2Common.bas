Attribute VB_Name = "W2Common"
Option Explicit
  Public RecNum As Long
  Public TaxText(1 To 10) As String * 2
  Public Emp2Rec(1) As EmpData2Type
  Public EHandle As Integer
  Public TRHandle As Integer
  Public SplitFlag As Boolean
  Public Const Manual = 2
  Public Const Normal = 1
  Public EntryType As Integer
  Public ScreenW As Long
  Public coladj As Double
  Public doAlign As Boolean
  Public CancelDoAlign As Boolean
  Public alnRpt$
  Public InFileNames(1 To 20) As String '7/20
  Public OutFileNames(1 To 20) As String '7/20
  Public ComputerName As String '7/20
  Public StartPath As String
  Public RegExit As Boolean
  Public RptOpt As Integer
  Public Twiddle As String
  
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long) '7/20
  
            Public Const PRData = "prdata\"
       Public Const EICFileName = "PREICTBL.DAT"
      Public Const UnitFileName = "PRUNIT.DAT"
       Public Const SysFileName = "PRSYS.DAT"
 Public Const TransWorkFileName = "PRTRANST.DAT"
 Public Const TransHistFileName = "PRTRANSH.DAT"
      Public Const EmpData1Name = "PREMP1.DAT"
      Public Const EmpData2Name = "PREMP2.DAT"
      Public Const EmpData3Name = "PREMP3.DAT"
       Public Const EmpIdxLName = "PREMPL.IDX"  'name idx
       Public Const EmpIdxNName = "PREMPN.IDX"  'numb idx
    Public Const ChecksFileName = "PRCHECKS.DAT"
 Public Const PPDefaultFileName = "PRPPDEF.DAT"
   Public Const DedCodeFileName = "PRDEDCOD.DAT"
   Public Const EmpDataFileMask = "PRRPTS\PREMPRPT.DPM"
  Public Const PrinterSetUpFile = "PRPRNSET.DAT"
    Public Const GLAcctIdxFile = "BAACCTDX.DAT"
    Public Const JGLAcctIdxFile = "GLACCT.IDX"
      Public Const AcctFileName = "GLACCT.DAT"
      Public Const TransFileName = "GLTRANS.DAT"
    Public Const MatCodeFileName = "PRDATA\PRMATCOD.DAT"
        Public Const W2SetupFile = "PRDATA\PRW2SETU.DAT"
         Public Const W2InfoFile = "PRDATA\PRW2INFO.DAT"
        Public Const W2PrintFile = "PRDATA\W2PRNT.DAT"
        Public Const W2RePrintFile = "PRDATA\W2REPRNT.DAT"
        Public Const W2ReprintIdx = "PRDATA\W2RPNIDX.DAT"
        Public Const W2ESubRA = "PRDATA\W2ESUBRA.DAT"
        Public Const W2ESubRE = "PRDATA\W2ESUBRE.DAT"
        Public Const W2ESubRW = "PRDATA\W2ESUBRW.DAT"
        Public Const W2ESubRT = "PRDATA\W2ESUBRT.DAT"
        Public Const W2ESubRF = "PRDATA\W2ESUBRF.DAT"
        Public Const W2ESubRU = "PRDATA\W2ESUBRU.DAT"
        Public Const W2ESubRO = "PRDATA\W2ESUBRO.DAT"
        Public Const W3InfoFile = "PRDATA\PRW3INFO.DAT"
    Public Const FederalTaxFileName = "PRFEDTAX.DAT"
Public Sub OpenW3Info(W3InfoHandle As Integer)
  Dim W3InfoRec As W3FormType
  Dim W3InfoRecLen As Integer
  W3InfoRecLen = Len(W3InfoRec)
  W3InfoHandle = FreeFile
  Open W3InfoFile For Random Shared As W3InfoHandle Len = W3InfoRecLen
End Sub
Public Sub OpenFedTaxFile(FedTaxFileHandle As Integer)
  Dim FedTaxFileRec As FederalTaxRecType
  Dim FedTaxRecLen As Integer
  FedTaxRecLen = Len(FedTaxFileRec)
  FedTaxFileHandle = FreeFile
  Open PRData + FederalTaxFileName For Random Shared As FedTaxFileHandle Len = FedTaxRecLen
End Sub
Public Sub OpenW2ESubRU(W2ESubRUHandle As Integer)
  Dim W2ESubRURec As W2ElectronicSubRU
  Dim W2ESubRURecLen As Integer
  W2ESubRURecLen = Len(W2ESubRURec)
  W2ESubRUHandle = FreeFile
  Open W2ESubRU For Random Shared As W2ESubRUHandle Len = W2ESubRURecLen
End Sub
Public Sub OpenW2ESubRO(W2ESubROHandle As Integer)
  Dim W2ESubRORec As W2ElectronicSubRO
  Dim W2ESubRORecLen As Integer
  W2ESubRORecLen = Len(W2ESubRORec)
  W2ESubROHandle = FreeFile
  Open W2ESubRO For Random Shared As W2ESubROHandle Len = W2ESubRORecLen
End Sub
Public Sub OpenW2ESubRA(W2ESubRAHandle As Integer)
  Dim W2ESubRARec As W2ElectronicSubRA
  Dim W2ESubRARecLen As Integer
  W2ESubRARecLen = Len(W2ESubRARec)
  W2ESubRAHandle = FreeFile
  Open W2ESubRA For Random Shared As W2ESubRAHandle Len = W2ESubRARecLen
End Sub
Public Sub OpenW2ESubRE(W2ESubREHandle As Integer)
  Dim W2ESubRERec As W2ElectronicSubRE
  Dim W2ESubRERecLen As Integer
  W2ESubRERecLen = Len(W2ESubRERec)
  W2ESubREHandle = FreeFile
  Open W2ESubRE For Random Shared As W2ESubREHandle Len = W2ESubRERecLen
End Sub
Public Sub OpenW2ESubRW(W2ESubRWHandle As Integer)
  Dim W2ESubRWRec As W2ElectronicSubRW
  Dim W2ESubRWRecLen As Integer
  W2ESubRWRecLen = Len(W2ESubRWRec)
  W2ESubRWHandle = FreeFile
  Open W2ESubRW For Random Shared As W2ESubRWHandle Len = W2ESubRWRecLen
End Sub
Public Sub OpenW2ESubRT(W2ESubRTHandle As Integer)
  Dim W2ESubRTRec As W2ElectronicSubRT
  Dim W2ESubRTRecLen As Integer
  W2ESubRTRecLen = Len(W2ESubRTRec)
  W2ESubRTHandle = FreeFile
  Open W2ESubRT For Random Shared As W2ESubRTHandle Len = W2ESubRTRecLen
End Sub
Public Sub OpenW2ESubRF(W2ESubRFHandle As Integer)
  Dim W2ESubRFRec As W2ElectronicSubRF
  Dim W2ESubRFRecLen As Integer
  W2ESubRFRecLen = Len(W2ESubRFRec)
  W2ESubRFHandle = FreeFile
  Open W2ESubRF For Random Shared As W2ESubRFHandle Len = W2ESubRFRecLen
End Sub
        
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
  frmW2Processing.Show
  DoEvents
End Sub
Public Sub OpenW2ReprintIdx(W2IdxHandle As Integer)
  Dim W2IdxRec As W2ReprintIdxType
  Dim W2IdxRecLen As Integer
  W2IdxRecLen = Len(W2IdxRec)
  W2IdxHandle = FreeFile
  Open W2ReprintIdx For Random Shared As W2IdxHandle Len = W2IdxRecLen
End Sub

Public Sub OpenW2Info(W2InfoHandle As Integer)
  Dim W2InfoRec As W2FormType
  Dim W2InfoRecLen As Integer
  W2InfoRecLen = Len(W2InfoRec)
  W2InfoHandle = FreeFile
  Open W2InfoFile For Random Shared As W2InfoHandle Len = W2InfoRecLen
  
End Sub
Public Sub OpenW2SetUp(W2SetUpHandle As Integer)
  Dim W2SetUpRec As W2SetUpType
  Dim W2SetUpRecLen As Integer
  W2SetUpRecLen = Len(W2SetUpRec)
  W2SetUpHandle = FreeFile
  Open W2SetupFile For Random Shared As W2SetUpHandle Len = W2SetUpRecLen
End Sub
      
Public Sub OpenPPDefaultFile(PPDefaultHandle As Integer)
  Dim PPDefaultRec As PeriodDefaultRecType
  Dim PPDefaultRecLen As Integer
  PPDefaultRecLen = Len(PPDefaultRec)
  PPDefaultHandle = FreeFile
  Open PRData + PPDefaultFileName For Random Shared As PPDefaultHandle Len = PPDefaultRecLen
End Sub
  
Public Sub OpenTransWorkFile(TransWorkFileHandle As Integer)
  Dim TransWorkFileRec As TransRecType
  Dim TransWorkRecLen As Integer
  TransWorkRecLen = Len(TransWorkFileRec)
  TransWorkFileHandle = FreeFile
  Open PRData + TransWorkFileName For Random Shared As TransWorkFileHandle Len = TransWorkRecLen
End Sub

Public Sub OpenTransHistFile(TransHistFileHandle As Integer)
  Dim TransHistFileRec As TransRecType
  Dim TransHistRecLen As Integer
  TransHistRecLen = Len(TransHistFileRec)
  TransHistFileHandle = FreeFile
  Open PRData + TransHistFileName For Random Shared As TransHistFileHandle Len = TransHistRecLen
End Sub
'****************************************************************************
'OldRounds a double precision value to nearest hundredth
'****************************************************************************
Public Function OldRound#(n As Double)
  OldRound# = Int(n * 100 + 0.5) / 100
'  OldRound# = Round(n, 2)
  'sofar mo6 betta
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
   
Public Sub OpenEmpData1File(EmpData1FileHandle As Integer)
  Dim EmpData1FileRec As EmpData1Type
  Dim EmpData1RecLen As Integer
  EmpData1RecLen = Len(EmpData1FileRec)
  EmpData1FileHandle = FreeFile
  Open PRData + EmpData1Name For Random Shared As EmpData1FileHandle Len = EmpData1RecLen
End Sub
Public Sub OpenDedCodeFile(DedCodeFileHandle As Integer)
  Dim DedCodeFileRec As DedCodeRecType
  Dim DedCodeRecLen As Integer
  DedCodeRecLen = Len(DedCodeFileRec)
  DedCodeFileHandle = FreeFile
  Open PRData + DedCodeFileName For Random Shared As DedCodeFileHandle Len = DedCodeRecLen
End Sub

Public Sub OpenEICFile(EICFileHandle As Integer)
  Dim EICFileRec As EICRecType
  Dim EICRecLen As Integer
  EICRecLen = Len(EICFileRec)
  EICFileHandle = FreeFile
  Open PRData + EICFileName For Random Shared As EICFileHandle Len = EICRecLen
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
      cnt = cnt + BigLen - 1
    Else
      NewText = NewText + thischar
    End If
  Next
  ReplaceString$ = Trim$(NewText)
  Text = ReplaceString$
End Function
Public Sub KillFile(FileName As String)
  If Exist(FileName$) Then
    Kill FileName$
  End If
End Sub

Public Function PromptSaveChanges(frm As Form) As SaveChangeOptions1
  frmChangedWarning.Show vbModal, frm
  PromptSaveChanges = frmChangedWarning.Selection
  Unload frmChangedWarning
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
  On Local Error Resume Next
  
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

Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
  
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

Public Function DirExists(ByVal strDirName As String) As Boolean
  On Error Resume Next
  
  Dim strFileName As String

  strFileName = strDirName & "\Nul"

  If (FileExists(strFileName)) Then
    DirExists = True
  Else
    DirExists = False
  End If
End Function

Public Sub W2ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
   frmLoadingW2Rpt.Show
   frmW2ViewPrint.ReportName = ReportFile$
   frmW2ViewPrint.Caption = Title
   frmW2ViewPrint.PgNum = PgNum
   frmW2ViewPrint.cmdAlignment.Visible = False
   If ForceSBar Then
     frmW2ViewPrint.fpMemo1.ScrollBars = BothFixed
   Else
     frmW2ViewPrint.fpMemo1.ScrollBars = BothAuto
   End If
   If Algn Then
     frmW2ViewPrint.cmdAlignment.Visible = True
     frmW2ViewPrint.AlignRpt = AlgnRptfile$
    End If
   frmW2ViewPrint.Show 1
   Unload frmLoadingW2Rpt
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
Public Function AddDashToSSN(ByVal SSN)
  Dim NewSSN As String
  If Mid(SSN, 4, 1) <> "-" And Mid(SSN, 7, 1) <> "-" Then
      NewSSN = Mid(SSN, 1, 3) + "-" + Mid(SSN, 5, 2) + "-" + Mid(SSN, 6, 4)
      AddDashToSSN = NewSSN
  Else
      AddDashToSSN = SSN
  End If
End Function

Public Function MonthName$(ByVal MonthNo As Integer)
  Select Case MonthNo
  Case 1
    MonthName$ = "January"
  Case 2
    MonthName$ = "February"
  Case 3
    MonthName$ = "March"
  Case 4
    MonthName$ = "April"
  Case 5
    MonthName$ = "May"
  Case 6
    MonthName$ = "June"
  Case 7
    MonthName$ = "July"
  Case 8
    MonthName$ = "August"
  Case 9
    MonthName$ = "September"
  Case 10
    MonthName$ = "October"
  Case 11
    MonthName$ = "November"
  Case 12
    MonthName$ = "December"
  End Select
  
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

Function FixDateSuffix(Source As String)

  Dim PDRRec As PeriodDefaultRecType
  Dim PHandle As Integer
  Dim EndDate As String
  Dim EDLen As String
  Dim TwoDigitYear As String
  Dim SourceLen As Integer
  
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PDRRec
  Close PHandle
  SourceLen = Len(Source)
  
  EndDate = MakeRegDate(PDRRec.PEREND)
  EDLen = Len(QPTrim$(EndDate))
  TwoDigitYear = Mid(EndDate, EDLen - 1, 2)
  FixDateSuffix = Mid(Source, 1, SourceLen - 2) + TwoDigitYear
  
End Function

Sub OldExtractW2Info(W2Type%, frm As Form)

  ReDim sp2(1) As String * 2
  ReDim W2SetUpRec(1) As W2SetUpType
  Dim W2Handle As Integer
  Dim W2SetUpRecLen As Integer
  Dim DedRec As DedCodeRecType
  Dim DedRecCnt As Integer
  Dim DedHandle As Integer
  Dim x As Integer
  Dim W2Year$
  Dim StrDate As Integer
  Dim EndDate As Integer
  Dim UnitRecLen As Integer
  Dim W2FormRecLen As Integer
  Dim Emp2RecLen As Integer
  Dim TranRecLen As Integer
  Dim ENumOfRec As Integer
  Dim TNumOfRec As Long
  Dim RptTitle$
  Dim UnitHandle As Integer
  Dim THandle As Integer
  Dim EHandle As Integer
  Dim WHandle As Integer
  Dim cnt As Long
  Dim ECnt As Integer
  Dim TotalTransRecs As Long
  Dim TCnt As Long
  Dim CntZZ As Integer
  Dim W2RWRec As W2ElectronicSubRW
  Dim BW2RWRec As W2ElectronicSubRW
  Dim RWHandle As Integer
  Dim FedTax As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FedTax
  Close FedTaxHandle
  FedSSMax = FedTax.FTMSSMW
  
  OpenDedCodeFile DedHandle
  DedRecCnt = LOF(DedHandle) / Len(DedRec)
  Close DedHandle
  
  ReDim Deds(0 To 50) As Double
  
  OpenW2SetUp W2Handle
  W2SetUpRec(1).ExtrYear = frm.fptxtYear.Text
'  W2SetUpRec(1).Deds(0).CHKDED = QPTrim$(frm.fpcomboRetire.Text)
  For x = 1 To DedRecCnt + 1
    frm.vaSpreadW2.Col = 2
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).CHKDED = QPTrim$(frm.vaSpreadW2.Text)
    frm.vaSpreadW2.Col = 3
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).AMTBOX = QPTrim$(frm.vaSpreadW2.Text)
    frm.vaSpreadW2.Col = 4
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).DedCode = QPTrim$(frm.vaSpreadW2.Text)
  Next x
  Put W2Handle, 1, W2SetUpRec(1)
  Close W2Handle

  W2Year = Val(frm.fptxtYear.Text)
  StrDate = Date2Num("01-01-" + frm.fptxtYear.Text)
  EndDate = Date2Num("12-31-" + frm.fptxtYear.Text)
  
  ReDim TranRec(1) As TransRecType
  ReDim UnitRec(1) As UnitFileRecType
  ReDim TPntr(0 To 200) As Long
  ReDim W2FormRec(1) As W2FormType
  ReDim BW2FormRec(1) As W2FormType
  
  UnitRecLen = Len(UnitRec(1))
  W2FormRecLen = Len(W2FormRec(1))
  Emp2RecLen = Len(Emp2Rec(1))
  TranRecLen = Len(TranRec(1))
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitRec(1)
  Close UnitHandle
  
  RptTitle$ = "Extracting W-2 Information"
  frmW2ShowPctComp.CmdCancel.Visible = False
  frmW2ShowPctComp.Label1 = "Extracting W-2 Information"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  OpenTransHistFile THandle
  TNumOfRec = LOF(THandle) \ Len(TranRec(1))

  'get trans action history pins
  ReDim TPins(1 To TNumOfRec) As Integer
  For cnt = 1 To TNumOfRec
    Get THandle, cnt, TranRec(1)
    TPins(cnt) = TranRec(1).EmpPin
  Next
  
  OpenEmpData2File EHandle
  ENumOfRec = LOF(EHandle) \ Len(Emp2Rec(1))
  
  OpenW2Info WHandle
  KillFile ("PRDATA\W2ESUBRW.DAT")
  KillFile ("PRDATA\W2ESUBRT.DAT")
  KillFile ("PRDATA\W2ESUBRO.DAT")
  KillFile ("PRDATA\W2ESUBRU.DAT")
  KillFile ("PRDATA\W2ESUBRF.DAT")
  
  OpenW2ESubRW RWHandle
  
  For ECnt = 1 To ENumOfRec
    Get EHandle, ECnt, Emp2Rec(1) '12/03
    GoSub GetEmpTranRecNums
    Select Case W2Type%
    Case 1
      If Emp2Rec(1).EMPSOCX <> "Y" Then
        GoSub CheckEmp
      Else
        W2FormRec(1) = BW2FormRec(1)
        W2RWRec = BW2RWRec
      End If
    Case 2
      If Emp2Rec(1).EMPSOCX = "Y" Then
        GoSub CheckEmp
      Else
        W2FormRec(1) = BW2FormRec(1)
        W2RWRec = BW2RWRec
      End If
    End Select
    
    Put WHandle, ECnt, W2FormRec(1)
    Put RWHandle, ECnt, W2RWRec '12/03
    
    frmW2ShowPctComp.ShowPctComp ECnt, ENumOfRec
'    Get RWHandle, ECnt, W2RWRec
'    W2RWRec.WageTips = W2RWRec.WageTips
  Next
  
  Unload frmW2ShowPctComp 'added 5/28/04
  EnableCloseButton frm.hwnd, True
  
  Close EHandle
  Close WHandle
  Close THandle
  Close RWHandle
  
  GoTo ExitW2SetUp
  
GetEmpTranRecNums:

  ReDim TPntr(0 To 1600)
  TotalTransRecs = 0
  For TCnt = 1 To TNumOfRec
    If TPins(TCnt) = Emp2Rec(1).EmpPin Then
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    End If
    TPntr(0) = TotalTransRecs
  Next
  Return

SumW2Info:
  W2FormRec(1) = BW2FormRec(1)
  W2RWRec = BW2RWRec
  For CntZZ = 0 To 50
    Deds(CntZZ) = 0
  Next

  For cnt = 1 To TPntr(0)
    Get THandle, TPntr(cnt), TranRec(1)
    If TranRec(1).CheckDate >= StrDate And TranRec(1).CheckDate <= EndDate Then
      W2FormRec(1).FEDWAGE = OldRound#(W2FormRec(1).FEDWAGE + TranRec(1).FedGrossPay)
      W2RWRec.WageTips = Currency2String(W2FormRec(1).FEDWAGE, 11) '12/03
      
      W2FormRec(1).FEDTAXWH = OldRound#(W2FormRec(1).FEDTAXWH + TranRec(1).FedTaxAmt)
      W2RWRec.FedTax = Currency2String(W2FormRec(1).FEDTAXWH, 11)
      
      W2FormRec(1).SOCWAGE = OldRound#(W2FormRec(1).SOCWAGE + TranRec(1).SocGrossPay)
      If W2FormRec(1).SOCWAGE > FedSSMax Then
        W2FormRec(1).SOCWAGE = FedSSMax
      End If
      W2RWRec.SSWages = Currency2String(W2FormRec(1).SOCWAGE, 11) '12/03
      W2FormRec(1).SOCTAXWH = OldRound#(W2FormRec(1).SOCTAXWH + TranRec(1).SocTaxAmt)
      W2RWRec.SSTax = Currency2String(W2FormRec(1).SOCTAXWH, 11) '12/03
      
      W2FormRec(1).MedWages = OldRound#(W2FormRec(1).MedWages + TranRec(1).MedGrossPay)
      W2RWRec.MedWages = Currency2String(W2FormRec(1).MedWages, 11) '12/03
      
      W2FormRec(1).MEDTAXWH = OldRound#(W2FormRec(1).MEDTAXWH + TranRec(1).MedTaxAmt)
      W2RWRec.MedTax = Currency2String(W2FormRec(1).MEDTAXWH, 11) '12/03
      
      W2FormRec(1).AdvEIC = OldRound#(W2FormRec(1).AdvEIC + TranRec(1).EICAmt)
      W2RWRec.AdvEIC = Currency2String(W2FormRec(1).AdvEIC, 11) '12/03
      
      W2FormRec(1).BENFBOX1 = OldRound#(W2FormRec(1).BENFBOX1 + TranRec(1).TaxFring)
      W2FormRec(1).State = UnitRec(1).UFSTATE
      W2FormRec(1).STAWAGE = OldRound#(W2FormRec(1).STAWAGE + TranRec(1).StaGrossPay)
      W2FormRec(1).STATAXWH = OldRound#(W2FormRec(1).STATAXWH + TranRec(1).StaTaxAmt)
      Deds(0) = OldRound#(Deds(0) + TranRec(1).RetireAmt)
      For CntZZ = 1 To 50
        Deds(CntZZ) = OldRound#(Deds(CntZZ) + TranRec(1).DAmt(CntZZ))
      Next
    End If
  Next
  W2RWRec.EmpFName = Emp2Rec(1).EmpFName
  W2RWRec.EmpMName = ""
  W2RWRec.EmpLName = Emp2Rec(1).EmpLName
  W2RWRec.EmpSuffix = ""
  Call ConvertNames(W2RWRec.EmpFName, W2RWRec.EmpMName, W2RWRec.EmpLName, W2RWRec.EmpSuffix)
  W2RWRec.EmpCity = QPTrim$(Emp2Rec(1).EmpCity) '12/03
  W2RWRec.EmpState = QPTrim$(Emp2Rec(1).EmpState) '12/03
  W2RWRec.EmpZip = QPTrim(Emp2Rec(1).EmpZip)
  W2RWRec.EmpZip = Mid(W2RWRec.EmpZip, 1, 5) '12/03
  W2RWRec.EmpZipX = QPTrim(Emp2Rec(1).EmpZip)
  W2RWRec.EmpZipX = Mid(W2RWRec.EmpZipX, 7, 4) '12/03
  W2RWRec.EmpSSN = ReplaceString(Emp2Rec(1).EmpSSN, "-", "") '12/03
  W2RWRec.EmpAdd1 = QPTrim$(Emp2Rec(1).EmpAddr1) '12/03
  W2RWRec.EmpAdd2 = QPTrim$(Emp2Rec(1).EMPADDR2) '12/03
  W2RWRec.SSTips = ""
  W2RWRec.DepCare = ""
  W2RWRec.NQPlan457 = ""
  W2RWRec.NQPNot457 = ""
  W2RWRec.StatuEmp = "0"
  W2RWRec.ThrdSckPay = "0"
  W2RWRec.ThrdSckAmt = "0"
  W2RWRec.RONum = "0"

Return

ApplyForm:
'*******************************
  If W2FormRec(1).BENFBOX1 > 0 Then 'BENFBOX1 = Taxable Fringe
    W2FormRec(1).FEDWAGE = OldRound#(W2FormRec(1).FEDWAGE + W2FormRec(1).BENFBOX1)
    W2RWRec.WageTips = Currency2String(W2FormRec(1).FEDWAGE, 11) '12/03
    
    W2FormRec(1).SOCWAGE = OldRound#(W2FormRec(1).SOCWAGE + W2FormRec(1).BENFBOX1)
    If W2FormRec(1).SOCWAGE > FedSSMax Then
      W2FormRec(1).SOCWAGE = FedSSMax
    End If
    W2RWRec.SSWages = Currency2String(W2FormRec(1).SOCWAGE, 11) '12/03
    
    W2FormRec(1).MedWages = OldRound#(W2FormRec(1).MedWages + W2FormRec(1).BENFBOX1)
    W2RWRec.MedWages = Currency2String(W2FormRec(1).MedWages, 11) '12/03
    W2FormRec(1).STAWAGE = OldRound#(W2FormRec(1).STAWAGE + W2FormRec(1).BENFBOX1)
  End If
'*******************************

  For CntZZ = 0 To 50
    If Len(QPTrim$(W2SetUpRec(1).Deds(CntZZ).CHKDED)) And Deds(CntZZ) > 0 Then
      Select Case Left$(W2SetUpRec(1).Deds(CntZZ).CHKDED, 1)
      Case "P" 'pension
        W2FormRec(1).BOX15c = "X"
        W2RWRec.RetPlan = "1" '12/03
      Case "D" 'deferred compensation
        W2FormRec(1).BOX15c = "X"
        W2RWRec.RetPlan = "1" '12/03
      End Select
    End If
  Next
  If W2RWRec.RetPlan <> "1" Then W2RWRec.RetPlan = "0"
'  If QPTrim$(Emp2Rec(1).EmpLName) = "CAVINESS" Then Stop
  
  W2FormRec(1).BOX13AMT = 0
  W2FormRec(1).BOX13AM1 = 0
  W2FormRec(1).BOX13AM2 = 0
  W2FormRec(1).BOX13AM3 = 0
  W2FormRec(1).BOX14AMT = 0
  W2FormRec(1).BOX14AM1 = 0
  
  For CntZZ = 0 To 50
    Select Case W2SetUpRec(1).Deds(CntZZ).AMTBOX
    Case "12a" '"13a"
      W2FormRec(1).BOX13AMT = OldRound#(W2FormRec(1).BOX13AMT + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TXt = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TXt) <> "" Then GoSub BOX13TXt
      End If
    Case "12b" '"13b"
      W2FormRec(1).BOX13AM1 = OldRound#(W2FormRec(1).BOX13AM1 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX1 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX1) <> "" Then GoSub BOX13TX1
      End If
    Case "12c" 'added Fall 04
      W2FormRec(1).BOX13AM2 = OldRound#(W2FormRec(1).BOX13AM2 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX2 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX2) <> "" Then GoSub BOX13TX2
      End If
    Case "12d" 'added Fall 04
      W2FormRec(1).BOX13AM3 = OldRound#(W2FormRec(1).BOX13AM3 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX3 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX3) <> "" Then GoSub BOX13TX3
      End If
    Case "14a"
      W2FormRec(1).BOX14AMT = OldRound#(W2FormRec(1).BOX14AMT + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX14TXT = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX14TXT) <> "" Then GoSub BOX14TXT
      End If
    Case "14b"
      W2FormRec(1).BOX14AM1 = OldRound#(W2FormRec(1).BOX14AM1 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX14TX1 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX14TX1) <> "" Then GoSub BOX14TX1
      End If
    End Select
  Next
  GoSub FinishIt
  Return
  
'  Exit Sub
  '12/03 from here to....
BOX13TXt: 'the W3 deferred compensation field only collects from Codes D-H and S
  Select Case QPTrim$(W2FormRec(1).BOX13TXt)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AMT, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT 'added fall 04
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AMT, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AMT, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AMT, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AMT, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "S" 'added 8/24/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
  End Select
  
Return

BOX13TX1:
  Select Case QPTrim$(W2FormRec(1).BOX13TX1)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM1, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM1, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM1, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM1, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM1, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "S"
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
  End Select
  Return

BOX13TX2:
  Select Case QPTrim$(W2FormRec(1).BOX13TX2)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM2, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM2, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM2, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM2, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM2, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "S"
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
  End Select
  Return

BOX13TX3:
  Select Case QPTrim$(W2FormRec(1).BOX13TX3)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM3, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM3, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM3, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM3, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM3, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "S"
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
  End Select
  Return

BOX14TXT:
  Select Case QPTrim$(W2FormRec(1).BOX14TXT)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX14AMT, 11)
  End Select
  Return

BOX14TX1:
  Select Case QPTrim$(W2FormRec(1).BOX14TX1)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX14AM1, 11)
  End Select
  Return
  
FinishIt:
  If Len(QPTrim$(W2RWRec.LifeIns)) = 0 Then
    W2RWRec.LifeIns = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr401k)) = 0 Then
    W2RWRec.Defr401k = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr403b)) = 0 Then
    W2RWRec.Defr403b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr408k6)) = 0 Then
    W2RWRec.Defr408k6 = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr457b)) = 0 Then
    W2RWRec.Defr457b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr501c18D)) = 0 Then
    W2RWRec.Defr501c18D = ""
  End If
  If Len(QPTrim$(W2RWRec.NonStaStcks)) = 0 Then
    W2RWRec.NonStaStcks = ""
  End If
  
  '...here
  Return

CheckEmp:
  If TPntr(0) Then            'if this emp has any transactions
    GoSub SumW2Info           'sum emp w2 info
    GoSub ApplyForm
  Else
    W2FormRec(1) = BW2FormRec(1)
    W2RWRec = BW2RWRec
  End If
Return

ExitW2SetUp:


End Sub

Public Sub W2ReportG()
  ReDim EmpCnt(1) As String * 26
  ReDim FedGross(1) As String * 13
  ReDim STAGROSS(1) As String * 13
  ReDim SocGross(1) As String * 13
  ReDim MedGross(1) As String * 13
'  ReDim Box13a(1) As String * 13
'  ReDim Box13b(1) As String * 13
'  ReDim Box13c(1) As String * 13
'  ReDim Box13d(1) As String * 13
  ReDim Box12a(1) As String * 13
  ReDim Box12b(1) As String * 13
  ReDim Box12c(1) As String * 13
  ReDim Box12d(1) As String * 13
  ReDim FedTax(1) As String * 13
  ReDim STATAX(1) As String * 13
  ReDim SocTax(1) As String * 13
  ReDim MedTax(1) As String * 13
  ReDim Box14a(1) As String * 13
  ReDim Box14b(1) As String * 13
  ReDim Box12(1) As String * 13
  ReDim AdvEIC(1) As String * 13
  ReDim Box10(1) As String * 13
  ReDim Box11(1) As String * 13

  ReDim Pg(1) As String * 3
  Dim Dash As String * 80
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  ReDim Line1(1) As String * 82
  Dim UHandle As Integer
  Dim Image1$
  Dim W2InfoRecLen As Integer
  Dim Emp2RecLen As Integer
  Dim RptName$
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim IdxNumOfRecs As Integer
  Dim RptTitle$
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim IdxLHandle As Integer
  Dim UTemp$
  Dim ECnt As Integer
  Dim TFedGross#, TStaGross#, TSocGross#, TMedGross#
  Dim TBox13a#, TBox13b#, TFedTax#, TStaTax#, TSocTax#
  Dim TBox13c#, TBox13d#, TBox12c#, TBox12d#
  Dim TMedTax#, TBox14a#, TBox14b#, TBox12#, TAdvEIC#
  Dim TBox10#, TBox11#, TBox12a#, TBox12b#
  Dim EmpOk As Boolean, dlm$
  
  On Error GoTo ErrorHandler
  dlm$ = "~"

  Image1$ = "##,###,##0.00"
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  W2InfoRecLen = Len(W2InfoRec(1))
  Emp2RecLen = Len(Emp2Rec(1))

  RptName$ = "PRRPTS\W2REPORTG.RPT"

  OpenEmpIdxLNameFile IdxLHandle
  IdxRecLen = 2
  IdxNumOfRecs = LOF(IdxLHandle) \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For cnt = 1 To IdxNumOfRecs
    Get IdxLHandle, cnt, IdxBuff(cnt)
  Next cnt
  Close IdxLHandle
  RptTitle$ = "W-2 Information Report"

  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  For cnt = 1 To IdxNumOfRecs
    Get WHandle, CLng(IdxBuff(cnt).DataRecNum), W2InfoRec(1)
    Get EHandle, CLng(IdxBuff(cnt).DataRecNum), Emp2Rec(1)
'    If QPTrim$(Emp2Rec(1).EmpLName) = "DEAN" Then Stop
    If Not Emp2Rec(1).Deleted Then
      GoSub PrintW2Data
    End If
  Next
  
  Close WHandle
  Close RHandle
  Close EHandle
  arW2Report.Show
  frmLoadingRpt.Show
  
  Exit Sub

PrintW2Data:

  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo SkipDontPrintEm
          End If
        End If
      End If
    End If
  End If
  
  ECnt = ECnt + 1
  TFedGross# = OldRound#(TFedGross# + W2InfoRec(1).FEDWAGE)
  TStaGross# = OldRound#(TStaGross# + W2InfoRec(1).STAWAGE)
  TSocGross# = OldRound#(TSocGross# + W2InfoRec(1).SOCWAGE)
  TMedGross# = OldRound#(TMedGross# + W2InfoRec(1).MedWages)
  TBox13a# = OldRound#(TBox13a# + W2InfoRec(1).BOX13AMT)
  TBox13b# = OldRound#(TBox13b# + W2InfoRec(1).BOX13AM1)
  TBox13c# = OldRound#(TBox13c# + W2InfoRec(1).BOX13AM2) 'added  fall 04
  TBox13d# = OldRound#(TBox13d# + W2InfoRec(1).BOX13AM3) 'added  fall 04
  
  TBox12a# = OldRound#(TBox13a# + W2InfoRec(1).BOX13AMT)
  TBox12b# = OldRound#(TBox13b# + W2InfoRec(1).BOX13AM1)
  TBox12c# = OldRound#(TBox13c# + W2InfoRec(1).BOX13AM2) 'added  fall 04
  TBox12d# = OldRound#(TBox13d# + W2InfoRec(1).BOX13AM3) 'added  fall 04
  
  TFedTax# = OldRound#(TFedTax# + W2InfoRec(1).FEDTAXWH)
  TStaTax# = OldRound#(TStaTax# + W2InfoRec(1).STATAXWH)
  TSocTax# = OldRound#(TSocTax# + W2InfoRec(1).SOCTAXWH)
  TMedTax# = OldRound#(TMedTax# + W2InfoRec(1).MEDTAXWH)
  TBox14a# = OldRound#(TBox14a# + W2InfoRec(1).BOX14AMT)
  TBox14b# = OldRound#(TBox14b# + W2InfoRec(1).BOX14AM1)
  TBox12# = OldRound#(TBox12# + W2InfoRec(1).BENFBOX1)
  TAdvEIC# = OldRound#(TAdvEIC# + W2InfoRec(1).AdvEIC)
  TBox10# = OldRound#(TBox10# + W2InfoRec(1).DEPNDCAR)
  TBox11# = OldRound#(TBox11# + W2InfoRec(1).NQPLAN)

  RSet FedGross(1) = Using$(Image1$, W2InfoRec(1).FEDWAGE)
  RSet STAGROSS(1) = Using$(Image1$, W2InfoRec(1).STAWAGE)
  RSet SocGross(1) = Using$(Image1$, W2InfoRec(1).SOCWAGE)
  RSet MedGross(1) = Using$(Image1$, W2InfoRec(1).MedWages)
  'box 13 is no longer used as of 2002
'  RSet Box13a(1) = Using$(Image1$, W2InfoRec(1).BOX13AMT)
'  RSet Box13b(1) = Using$(Image1$, W2InfoRec(1).BOX13AM1)
  RSet Box12a(1) = Using$(Image1$, W2InfoRec(1).BOX13AMT)
  RSet Box12b(1) = Using$(Image1$, W2InfoRec(1).BOX13AM1)
  RSet Box12c(1) = Using$(Image1$, W2InfoRec(1).BOX13AM2)
  RSet Box12d(1) = Using$(Image1$, W2InfoRec(1).BOX13AM3)
  
  '8/25/04 insert 12c and 12d
  RSet FedTax(1) = Using$(Image1$, W2InfoRec(1).FEDTAXWH)
  RSet STATAX(1) = Using$(Image1$, W2InfoRec(1).STATAXWH)
  RSet SocTax(1) = Using$(Image1$, W2InfoRec(1).SOCTAXWH)
  RSet MedTax(1) = Using$(Image1$, W2InfoRec(1).MEDTAXWH)
  RSet Box14a(1) = Using$(Image1$, W2InfoRec(1).BOX14AMT)
  RSet Box14b(1) = Using$(Image1$, W2InfoRec(1).BOX14AM1)
  RSet AdvEIC(1) = Using$(Image1$, W2InfoRec(1).AdvEIC)
  RSet Box10(1) = Using$(Image1$, W2InfoRec(1).DEPNDCAR)
  RSet Box11(1) = Using$(Image1$, W2InfoRec(1).NQPLAN)
  RSet Box12(1) = Using$(Image1$, W2InfoRec(1).BENFBOX1)

  '                     0                1                          2
  Print #RHandle, Unit(1).UFEMPR; dlm; Date$; dlm; UCase$(QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName)); dlm;
  '                               3                                         4
  Print #RHandle, Using$(Image1$, W2InfoRec(1).AdvEIC); dlm; Using$(Image1$, W2InfoRec(1).DEPNDCAR); dlm;
  '                          non-qualified plan  5                             6
  Print #RHandle, Using$(Image1$, W2InfoRec(1).NQPLAN); dlm; Using$(Image1$, W2InfoRec(1).BENFBOX1); dlm;
  '                   7                 8                9               10
  Print #RHandle, FedGross(1); dlm; STAGROSS(1); dlm; SocGross(1); dlm; MedGross(1); dlm;
  '                  11              12              13              14              15
  Print #RHandle, Box12a(1); dlm; Box12b(1); dlm; FedTax(1); dlm; STATAX(1); dlm; SocTax(1); dlm;
  '                  16              17             18
  Print #RHandle, MedTax(1); dlm; Box14a(1); dlm; Box14b(1); dlm; Box12c(1); dlm; Box12d(1)
SkipDontPrintEm:


Return

ErrorHandler:
   If Err.Number = 6 Then
     frmW2Message.Label1.Caption = "Error: some of the currency values attempting to be loaded are causing overflow problems. This is probably because extraction has not taken place. Please try extracting the latest data before continuing."
     frmW2Message.Label1.Top = 650
     frmW2Message.Show vbModal
     MainLog ("Error: User warned that during the printing for Form W2 (graphics) report some of the values are causing overflow issues and that extraction needed to be processed.")
     Close
     Exit Sub
   End If
   
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "W2Common", "W2ReportG", Erl)
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

End Sub

Sub W2PrintTForms(frm As Form)
  ReDim PEMPCity(1) As String * 20
  ReDim PEmpSSN(1) As String * 15
  ReDim BTxt14(1) As String * 5
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  Dim Image1$, Image2$, Image$
  Dim UHandle As Integer
  Dim CtrlNumb&
  Dim DidOne As Integer
  Dim RptName$, RptTitle$
  Dim PrnCnt As Integer
  Dim MaxPrn As Integer
  Dim IdxNumOfRecs As Integer
  Dim IdxLHandle As Integer
  Dim x As Integer
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim SocTip$, AlocTip$, AdvEicP$, DepCare$, NQP$
  Dim FEDWAGE#, FEDTAXWH#, SOCWAGE#, SOCTAXWH#
  Dim MedWages#, MEDTAXWH#, SocTips#, ALLOCTIP#
  Dim AdvEIC#, DEPNDCAR#, NQPLAN#, BENFBOX1#
  Dim BOX13AMT#, BOX13AM1#, STATAXWH#, STAWAGE#
  Dim BOX13AM2#, BOX13AM3#
  Dim Box13Amt1$, Box14Amt1$
  Dim StateWage$, StateTax$
  Dim W2IdxRec As W2ReprintIdxType
  Dim IdxRHandle As Integer
  Dim ThisIdx As Long
  Dim StateID As String
  
  On Error GoTo ErrorHandler
  
  ThisIdx = 0
  If QPTrim$(frm.fptxtStartConNum.Text = "") Then
    MsgBox "Please make an entry in the Starting Control Number field"
    frm.fptxtStartConNum.SetFocus
    Exit Sub
  End If
  CtrlNumb& = CLng(frm.fptxtStartConNum.Text)
  Image1$ = "######.##"
  Image2$ = "#####.##"
  Image$ = "######"
  OpenUnitFile UHandle
  Get UHandle, , Unit(1)
  Close UHandle
  StateID = ""
  If Unit(1).UFSTATE = "VA" Then
   StateID = Mid(Unit(1).UFSTAID, 1, 2)
   StateID = StateID + QPTrim$(Unit(1).UFFEDID)
   StateID = StateID + Mid(QPTrim$(Unit(1).UFSTAID), 3, Len(QPTrim$(Unit(1).UFSTAID)))
  End If

  RptName$ = W2PrintFile
  RptTitle$ = "W-2 Forms Printing"

  PrnCnt = 0
  MaxPrn = 41
  OpenEmpIdxLNameFile IdxLHandle
  IdxNumOfRecs = LOF(IdxLHandle) \ 2
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For x = 1 To IdxNumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  Close IdxLHandle

  frmW2ShowPctComp.Label1 = "Printing W2 Forms"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  KillFile "PRDATA\W2RPNIDX.DAT"
  OpenW2ReprintIdx IdxRHandle
  
  For cnt = 1 To IdxNumOfRecs
    Get WHandle, CLng(IdxBuff(cnt).DataRecNum), W2InfoRec(1)
    Get EHandle, CLng(IdxBuff(cnt).DataRecNum), Emp2Rec(1)
    GoSub PrintW2Form
  frmW2ShowPctComp.ShowPctComp cnt, IdxNumOfRecs
  If frmW2ShowPctComp.Out = True Then
    Close
    frmW2ShowPctComp.Out = False
    frm.cmdEscape.Enabled = True
    frm.cmdProcess.Enabled = True
    EnableCloseButton frm.hwnd, True
    Unload frmW2ShowPctComp
  End If
  Next
  Close IdxRHandle
  Close RHandle
  Close WHandle
  Close EHandle

  W2ViewPrint RptName$, RptTitle$, True, , False
Exit Sub

PrintW2Form:
  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo DontPrintEm
          End If
        End If
      End If
    End If
  End If

  DidOne = DidOne + 1

  LSet PEMPCity(1) = Emp2Rec(1).EmpCity
  LSet PEmpSSN(1) = Left$(Emp2Rec(1).EmpSSN, 3) + "-" + Mid$(Emp2Rec(1).EmpSSN, 4, 2) + "-" + Right$(QPTrim$(Emp2Rec(1).EmpSSN), 4)
'start of w2 forms printing
  Print #RHandle, "!" '1
  Print #RHandle,     '2
  Print #RHandle,     '3 'added 10/13/04 for 2004 forms
  Print #RHandle,     '4
'  Print #RHandle, Tab(6); Using(Image$, CtrlNumb&) '5
  Print #RHandle, Tab(23); PEmpSSN(1) '5
  Print #RHandle,     '6
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFFEDID); Tab(51); Using(Image1$, W2InfoRec(1).FEDWAGE); Tab(68); Using(Image1$, W2InfoRec(1).FEDTAXWH) '7
  Print #RHandle, '8
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFEMPR); Tab(51); Using(Image1$, W2InfoRec(1).SOCWAGE); Tab(68); Using(Image1$, W2InfoRec(1).SOCTAXWH) '9
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFADDR1) '10
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFADDR2); Tab(51); Using(Image1$, W2InfoRec(1).MedWages); Tab(68); Using(Image1$, W2InfoRec(1).MEDTAXWH) '11
  Print #RHandle, Tab(5); Left$(QPTrim$(Unit(1).UFCITY), 15) + " " + QPTrim$(Unit(1).UFSTATE) + " " + QPTrim$(Unit(1).UFZIP) '12
  
  If W2InfoRec(1).SocTips > 0 Then
    SocTip$ = Using(Image1$, W2InfoRec(1).SocTips)
  Else
    SocTip$ = Space$(Len(Image1$))
  End If
  If W2InfoRec(1).ALLOCTIP > 0 Then
    AlocTip$ = Using(Image1$, W2InfoRec(1).ALLOCTIP)
  Else
    AlocTip$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(51); SocTip$; Tab(68); AlocTip$ '13

  Print #RHandle,     ' line 14

  If W2InfoRec(1).AdvEIC > 0 Then
    AdvEicP$ = Using(Image1$, W2InfoRec(1).AdvEIC)
  Else
    AdvEicP$ = Space$(Len(Image1$))
  End If

  If W2InfoRec(1).DEPNDCAR > 0 Then
    DepCare$ = Using(Image1$, W2InfoRec(1).DEPNDCAR)
  Else
    DepCare$ = Space$(Len(Image1$))
  End If

'  Print #RHandle, Tab(5); PEmpSSN(1); Tab(51); AdvEicP$; Tab(68); DepCare$ '15
  Print #RHandle, Tab(15); Using(Image$, CtrlNumb&); Tab(51); AdvEicP$; Tab(68); DepCare$ '15
  Print #RHandle, '16

  If W2InfoRec(1).NQPLAN > 0 Then
    NQP$ = Using(Image1$, W2InfoRec(1).NQPLAN)
  Else
    NQP$ = Space$(Len(Image1$))
  End If
  
  If W2InfoRec(1).BOX13AMT > 0 Then    '12a
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AMT)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(5); QPTrim$(Emp2Rec(1).EmpFName); Tab(28); QPTrim$(Emp2Rec(1).EmpLName); Tab(51); NQP$;
'  Print #RHandle, Tab(66); Left$(QPTrim$(W2InfoRec(1).BOX13TXt) + Space$(1), 1); "   "; Box13Amt1$ '16'changed for 2004 forms
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TXt) + Space$(1), 1); "   "; Box13Amt1$ '17
  
  Print #RHandle, 'Tab(5); QPTrim$(Emp2Rec(1).EmpAddr1) '18

  If W2InfoRec(1).BOX13AM1 > 0 Then
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM1)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(5); QPTrim$(Emp2Rec(1).EmpAddr1);
'  Print #RHandle, Tab(49); QPTrim$(W2InfoRec(1).BOX15A); Tab(54); QPTrim$(W2InfoRec(1).BOX15c); Tab(59); QPTrim$(W2InfoRec(1).BOX15G);'changed for 2004 forms
  Print #RHandle, Tab(48); QPTrim$(W2InfoRec(1).BOX15A); Tab(53); QPTrim$(W2InfoRec(1).BOX15c); Tab(58); QPTrim$(W2InfoRec(1).BOX15G);
'  Print #RHandle, Tab(66); Left$(QPTrim$(W2InfoRec(1).BOX13TX1) + Space$(1), 1); "   "; Box13Amt1$ '18'changed for 2004 forms
  Print #RHandle, Tab(63); Left$(QPTrim$(W2InfoRec(1).BOX13TX1) + Space$(1), 2); "   "; Box13Amt1$ '19

  '8/25/04 inserted 12c and 12d

  Print #RHandle, Tab(5); QPTrim$(Emp2Rec(1).EMPADDR2) '20

  If W2InfoRec(1).BOX13AM2 > 0 Then 'added fall 04
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM2)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If
  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TXT) + Space$(5)        'line 14 a
  If W2InfoRec(1).BOX14AMT > 0 Then
    Box14Amt1$ = Using(Image1$, W2InfoRec(1).BOX14AMT)
  Else
    Box14Amt1$ = Space$(Len(Image1$))
  End If
  Print #RHandle, Tab(5); PEMPCity(1) + " "; QPTrim$(Emp2Rec(1).EmpState); " "; QPTrim$(Emp2Rec(1).EmpZip); Tab(47); BTxt14(1); Box14Amt1$; '21
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TX2) + Space$(1), 1); "   "; Box13Amt1$ '21'added fall 04

  If W2InfoRec(1).BOX13AM3 > 0 Then 'added fall 04
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM3)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If
  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TX1) + Space$(5)        'line 14 b
  If W2InfoRec(1).BOX14AM1 > 0 Then
    Box14Amt1$ = Using(Image1$, W2InfoRec(1).BOX14AM1)
  Else
    Box14Amt1$ = Space$(Len(Image1$))
  End If
  Print #RHandle, Tab(47); BTxt14(1); Box14Amt1$ '22
'  Print #RHandle, '22 'commented on 8/25/04
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TX3) + Space$(1), 1); "   "; Box13Amt1$ '23'added fall 04
  Print #RHandle, '24

  If W2InfoRec(1).STAWAGE > 0 Then
    StateWage$ = Using(Image1$, W2InfoRec(1).STAWAGE)
  Else
    StateWage$ = Space$(Len(Image1$))
  End If

  If W2InfoRec(1).STATAXWH > 0 Then
    StateTax$ = Using(Image1$, W2InfoRec(1).STATAXWH)
  Else
    StateTax$ = Space$(Len(Image1$))
  End If

  Print #RHandle, '25
  If Unit(1).UFSTATE = "VA" Then
   Print #RHandle, Tab(5); W2InfoRec(1).State; Tab(11); StateID; Tab(28); StateWage$; Tab(39); StateTax$ '26
  Else
'  Print #RHandle, Tab(3); W2InfoRec(1).State; Tab(9); QPTrim$(Unit(1).UFSTAID); Tab(26); StateWage$; Tab(38); StateTax$ '25'omitted for 2004 forms
    Print #RHandle, Tab(5); W2InfoRec(1).State; Tab(11); QPTrim$(Unit(1).UFSTAID); Tab(28); StateWage$; Tab(39); StateTax$ '26
  End If
  Print #RHandle, '27
  Print #RHandle, '28
  Print #RHandle, '29
  Print #RHandle, '30
  Print #RHandle, '31
  Print #RHandle, '32
'  Print #RHandle, '32 'omitted for 2004 forms
  Print #RHandle, "!" '33

  PrnCnt = PrnCnt + 1
  ThisIdx = ThisIdx + 1
  W2IdxRec.RECNO = CLng(IdxBuff(cnt).DataRecNum)
  W2IdxRec.CONTNUM = CtrlNumb&
  Put IdxRHandle, ThisIdx, W2IdxRec
  CtrlNumb& = CtrlNumb& + 1
DontPrintEm:
  Return

SumW2SubTotal:
  FEDWAGE# = OldRound#(FEDWAGE# + W2InfoRec(1).FEDWAGE)
  FEDTAXWH# = OldRound#(FEDTAXWH# + W2InfoRec(1).FEDTAXWH)
  SOCWAGE# = OldRound#(SOCWAGE# + W2InfoRec(1).SOCWAGE)
  SOCTAXWH# = OldRound#(SOCTAXWH# + W2InfoRec(1).SOCTAXWH)
  MedWages# = OldRound#(MedWages# + W2InfoRec(1).MedWages)
  MEDTAXWH# = OldRound#(MEDTAXWH# + W2InfoRec(1).MEDTAXWH)
  SocTips# = OldRound#(SocTips# + W2InfoRec(1).SocTips)
  ALLOCTIP# = OldRound#(ALLOCTIP# + W2InfoRec(1).ALLOCTIP)
  AdvEIC# = OldRound#(AdvEIC# + W2InfoRec(1).AdvEIC)
  DEPNDCAR# = OldRound#(DEPNDCAR# + W2InfoRec(1).DEPNDCAR)
  NQPLAN# = OldRound#(NQPLAN# + W2InfoRec(1).NQPLAN)
  '11
  BENFBOX1# = OldRound#(BENFBOX1# + W2InfoRec(1).BENFBOX1)
  BOX13AMT# = OldRound#(BOX13AMT# + W2InfoRec(1).BOX13AMT)
  BOX13AM1# = OldRound#(BOX13AM1# + W2InfoRec(1).BOX13AM1)
  BOX13AM2# = OldRound#(BOX13AM2# + W2InfoRec(1).BOX13AM2)
  BOX13AM3# = OldRound#(BOX13AM3# + W2InfoRec(1).BOX13AM3)
  STATAXWH# = OldRound#(STATAXWH# + W2InfoRec(1).STATAXWH)
  STAWAGE# = OldRound#(STAWAGE# + W2InfoRec(1).STAWAGE)
Return

ErrorHandler:
   Unload frmW2ShowPctComp
   If Err.Number = 6 Then
     frmW2Message.Label1.Caption = "Error: some of the currency values are causing overflow problems. This is probably because extraction has not taken place. Please try extracting the latest data before continuing."
     frmW2Message.Label1.Top = 650
     frmW2Message.Show vbModal
     MainLog ("Error: User warned that during the printing for Form W2 Text some of the values are causing overflow issues and that extraction needed to be processed.")
     Close
     Exit Sub
   End If
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "W2Common", "W2PrintTForms", Erl)
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

End Sub

Public Sub UnloadAllFormsAndOpn(RegExit As Boolean)
  Unload frm914W2Setup
  Unload frmChangedWarning '8/13
  Unload frmEditReviewEmpW2
  Unload frmLoadingW2Rpt
  Unload frmMedW2Setup
  Unload frmW2EmpInfo
  Unload frmW2FormsPrinting
  Unload frmW2Print
  Unload frmW2Processing
  Unload frmW2ShowPctComp
  Unload frmW2ViewPrint
  Unload frmWarnW2FilesMissing '8/13
  Unload arW2Report
  Unload frmLoadingRpt
  Unload frmRestartPrinter
  If PWcnt = -3 Then Exit Sub
  If RegExit = False Then
    ClearInUse PWcnt
  End If
End Sub
'7/20 added this function
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
    frmWarnW2FilesMissing.Show vbModal, frm
    For x = 1 To ThisMany
      InFileNames(x) = ""
      OutFileNames(x) = ""
    Next x
  End If
End Function

Public Sub W2ReportT()
  ReDim EmpCnt(1) As String * 26
  ReDim FedGross(1) As String * 13
  ReDim STAGROSS(1) As String * 13
  ReDim SocGross(1) As String * 13
  ReDim MedGross(1) As String * 13
  ReDim Box13a(1) As String * 13
  ReDim Box13b(1) As String * 13
  ReDim Box13c(1) As String * 13
  ReDim Box13d(1) As String * 13
  ReDim FedTax(1) As String * 13
  ReDim STATAX(1) As String * 13
  ReDim SocTax(1) As String * 13
  ReDim MedTax(1) As String * 13
  ReDim Box14a(1) As String * 13
  ReDim Box14b(1) As String * 13
  ReDim Box12(1) As String * 13
  ReDim AdvEIC(1) As String * 13
  ReDim Box10(1) As String * 13
  ReDim Box11(1) As String * 13
  ReDim Pg(1) As String * 3
  Dim Dash As String * 80
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  ReDim Line1(1) As String * 82
  Dim UHandle As Integer
  Dim Image1$
  Dim W2InfoRecLen As Integer
  Dim Emp2RecLen As Integer
  Dim RptName$
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim IdxNumOfRecs As Integer
  Dim RptTitle$
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim IdxLHandle As Integer
  Dim FF$
  Dim UTemp$
  Dim ECnt As Integer
  Dim TFedGross#, TStaGross#, TSocGross#, TMedGross#
  Dim TBox13a#, TBox13b#, TFedTax#, TStaTax#, TSocTax#
  Dim TBox13c#, TBox13d#
  Dim TMedTax#, TBox14a#, TBox14b#, TBox12#, TAdvEIC#
  Dim TBox10#, TBox11#
  Dim EmpOk As Boolean
  
  On Error GoTo ErrorHandler
  
  FF$ = Chr$(12)
  
  Dash = String$(78, "-") ' + CrLf$

  Image1$ = "##,###,##0.00"
  OpenUnitFile UHandle
  Get UHandle, 1, Unit(1)
  Close UHandle
  W2InfoRecLen = Len(W2InfoRec(1))
  Emp2RecLen = Len(Emp2Rec(1))

  RptName$ = "PRRPTS\W2REPORT.RPT"

  LineCnt = 0
  MaxLines = 50
  Page = 0
  OpenEmpIdxLNameFile IdxLHandle
  IdxRecLen = 2
  IdxNumOfRecs = LOF(IdxLHandle) \ IdxRecLen
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For cnt = 1 To IdxNumOfRecs
    Get IdxLHandle, cnt, IdxBuff(cnt)
  Next cnt
  Close IdxLHandle
  RptTitle$ = "W-2 Information Report"

  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  GoSub PrintW2Header
  For cnt = 1 To IdxNumOfRecs
    Get WHandle, CLng(IdxBuff(cnt).DataRecNum), W2InfoRec(1)
    Get EHandle, CLng(IdxBuff(cnt).DataRecNum), Emp2Rec(1)
    If Not Emp2Rec(1).Deleted Then
      GoSub PrintW2Data
      If LineCnt >= MaxLines Then
        Print #RHandle, FF$
        GoSub PrintW2Header
      End If
    End If
  Next

  GoSub PrintW2Summary

  Close WHandle
  Close RHandle
  Close EHandle

  W2ViewPrint RptName$, RptTitle$, False, False, False, False
  Exit Sub

PrintW2Header:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(80)
  LSet UTemp$ = Unit(1).UFEMPR
  Mid$(UTemp$, 71) = "Page:" + Pg(1)

  Print #RHandle, UTemp$
  Print #RHandle, "W-2 Information Report"
  Print #RHandle, "Report Date: " + Date$
  Print #RHandle, "Employee Name  "
  Print #RHandle, "                   Adv EIC       Box 10       Box 11      Box 12a      Box 12b"
  Print #RHandle, "    Fed Gross  State Gross    Soc Gross    Med Gross      Box 12c      Box 12d"
  Print #RHandle, "      Fed Tax    State Tax      Soc Tax      Med Tax       Box 14       Box 14"
  Print #RHandle, Dash
  LineCnt = 8
  UTemp$ = ""
  Return


PrintW2Data:

  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo SkipDontPrintEm
          End If
        End If
      End If
    End If
  End If
  
  
  ECnt = ECnt + 1
  TFedGross# = OldRound#(TFedGross# + W2InfoRec(1).FEDWAGE)
  TStaGross# = OldRound#(TStaGross# + W2InfoRec(1).STAWAGE)
  TSocGross# = OldRound#(TSocGross# + W2InfoRec(1).SOCWAGE)
  TMedGross# = OldRound#(TMedGross# + W2InfoRec(1).MedWages)
  TBox13a# = OldRound#(TBox13a# + W2InfoRec(1).BOX13AMT)
  TBox13b# = OldRound#(TBox13b# + W2InfoRec(1).BOX13AM1)
  TBox13c# = OldRound#(TBox13c# + W2InfoRec(1).BOX13AM2) 'added fall 04
  TBox13d# = OldRound#(TBox13d# + W2InfoRec(1).BOX13AM3) 'added fall 04
  TFedTax# = OldRound#(TFedTax# + W2InfoRec(1).FEDTAXWH)
  TStaTax# = OldRound#(TStaTax# + W2InfoRec(1).STATAXWH)
  TSocTax# = OldRound#(TSocTax# + W2InfoRec(1).SOCTAXWH)
  TMedTax# = OldRound#(TMedTax# + W2InfoRec(1).MEDTAXWH)
  TBox14a# = OldRound#(TBox14a# + W2InfoRec(1).BOX14AMT)
  TBox14b# = OldRound#(TBox14b# + W2InfoRec(1).BOX14AM1)
  TBox12# = OldRound#(TBox12# + W2InfoRec(1).BENFBOX1)
  TAdvEIC# = OldRound#(TAdvEIC# + W2InfoRec(1).AdvEIC)
  TBox10# = OldRound#(TBox10# + W2InfoRec(1).DEPNDCAR)
  TBox11# = OldRound#(TBox11# + W2InfoRec(1).NQPLAN)

  RSet FedGross(1) = Using$(Image1$, W2InfoRec(1).FEDWAGE)
  RSet STAGROSS(1) = Using$(Image1$, W2InfoRec(1).STAWAGE)
  RSet SocGross(1) = Using$(Image1$, W2InfoRec(1).SOCWAGE)
  RSet MedGross(1) = Using$(Image1$, W2InfoRec(1).MedWages)

  RSet Box13a(1) = Using$(Image1$, W2InfoRec(1).BOX13AMT)
  RSet Box13b(1) = Using$(Image1$, W2InfoRec(1).BOX13AM1)
  RSet Box13c(1) = Using$(Image1$, W2InfoRec(1).BOX13AM2) 'added fall 04
  RSet Box13d(1) = Using$(Image1$, W2InfoRec(1).BOX13AM3) 'added fall 04
  RSet FedTax(1) = Using$(Image1$, W2InfoRec(1).FEDTAXWH)
  RSet STATAX(1) = Using$(Image1$, W2InfoRec(1).STATAXWH)
  RSet SocTax(1) = Using$(Image1$, W2InfoRec(1).SOCTAXWH)
  RSet MedTax(1) = Using$(Image1$, W2InfoRec(1).MEDTAXWH)
  RSet Box14a(1) = Using$(Image1$, W2InfoRec(1).BOX14AMT)
  RSet Box14b(1) = Using$(Image1$, W2InfoRec(1).BOX14AM1)
  RSet AdvEIC(1) = Using$(Image1$, W2InfoRec(1).AdvEIC)
  RSet Box10(1) = Using$(Image1$, W2InfoRec(1).DEPNDCAR)
  RSet Box11(1) = Using$(Image1$, W2InfoRec(1).NQPLAN)
  RSet Box12(1) = Using$(Image1$, W2InfoRec(1).BENFBOX1)

  LSet Line1(1) = UCase$(QPTrim$(Emp2Rec(1).EmpLName) + ", " + QPTrim$(Emp2Rec(1).EmpFName))
  Print #RHandle, Line1(1)
  Line1(1) = ""
  Mid$(Line1(1), 14) = AdvEIC(1) 'AdvEIC(1)
  Mid$(Line1(1), 27) = Box10(1)
  Mid$(Line1(1), 40) = Box11(1)
  Mid$(Line1(1), 53) = Box13a(1) 'Box12(1)..8/25/04 Box12(1) was not being printed in the old version so this slot is now being used with the fall 04 update
  Mid$(Line1(1), 66) = Box13b(1)
  '8/25/04 insert 12c and 12d
  Print #RHandle, Line1(1)
  Print #RHandle, FedGross(1) + STAGROSS(1) + SocGross(1) + MedGross(1) + Box13c(1) + Box13d(1)
  Print #RHandle, FedTax(1) + STATAX(1) + SocTax(1) + MedTax(1) + Box14a(1) + Box14b(1)
  Print #RHandle,
  LineCnt = LineCnt + 5
SkipDontPrintEm:
Return

PrintW2Summary:
  EmpOk = False
  RSet EmpCnt(1) = Using$("####", ECnt)
  RSet FedGross(1) = Using$(Image1$, TFedGross#)
  RSet STAGROSS(1) = Using$(Image1$, TStaGross#)
  RSet SocGross(1) = Using$(Image1$, TSocGross#)
  RSet MedGross(1) = Using$(Image1$, TMedGross#)
  RSet Box13a(1) = Using$(Image1$, TBox13a#)
  RSet Box13b(1) = Using$(Image1$, TBox13b#)
  RSet Box13c(1) = Using$(Image1$, TBox13c#) 'added 8/25/04
  RSet Box13d(1) = Using$(Image1$, TBox13d#) 'added 8/25/04
  RSet FedTax(1) = Using$(Image1$, TFedTax#)
  RSet STATAX(1) = Using$(Image1$, TStaTax#)
  RSet SocTax(1) = Using$(Image1$, TSocTax#)
  RSet MedTax(1) = Using$(Image1$, TMedTax#)

  RSet AdvEIC(1) = Using$(Image1$, TAdvEIC#)
  RSet Box12(1) = Using$(Image1$, TBox12#)
  RSet Box10(1) = Using$(Image1$, TBox10#)
  RSet Box11(1) = Using$(Image1$, TBox11#)

  RSet Box14a(1) = Using$(Image1$, TBox14a#)
  RSet Box14b(1) = Using$(Image1$, TBox14b#)

PrintEm: '8/25/04 insert 12c and 12d
  Print #RHandle, FF$
  Print #RHandle, Dash
  Print #RHandle, "Report Totals"
  Print #RHandle, "     W2 Forms      Adv EIC       Box 10       Box 11      Box 12a      Box 12b"
  
  Print #RHandle, Tab(10); Using$("####", ECnt) + AdvEIC(1) + Box10(1) + Box11(1) + Box13a(1) + Box13b(1)
'  Print #RHandle, EmpCnt(1) + AdvEIC(1) + Box10(1) + Box11(1) + Box12(1)
  Print #RHandle, "    Fed Gross  State Gross    Soc Gross    Med Gross      Box 12c      Box 12d"
  Print #RHandle, FedGross(1) + STAGROSS(1) + SocGross(1) + MedGross(1) + Box13c(1) + Box13d(1)
  Print #RHandle, "      Fed Tax    State Tax      Soc Tax      Med Tax       Box 14       Box 14"
  Print #RHandle, FedTax(1) + STATAX(1) + SocTax(1) + MedTax(1) + Box14a(1) + Box14b(1)
  Print #RHandle, Dash
  Print #RHandle, FF$

  Return
  
ErrorHandler:
   Unload frmW2ShowPctComp
   If Err.Number = 6 Then
     frmW2Message.Label1.Caption = "Error: some of the currency values are causing overflow problems. This is probably because extraction has not taken place. Please try extracting the latest data before continuing."
     frmW2Message.Label1.Top = 650
     frmW2Message.Show vbModal
     MainLog ("Error: User warned that during the printing for Form W2 text report some of the values are causing overflow issues and that extraction needed to be processed.")
     Close
     Exit Sub
   End If
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "W2Common", "W2ReportT", Erl)
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
  
End Sub

Sub W2PrintGForms(frm As Form)
  ReDim PEMPCity(1) As String * 20
  ReDim PEmpSSN(1) As String * 15
  ReDim BTxt14(1) As String * 5
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  Dim Image1$, Image2$, Image$
  Dim UHandle As Integer
  Dim CtrlNumb&
  Dim DidOne As Integer
  Dim RptName$, RptTitle$
  Dim PrnCnt As Integer
  Dim MaxPrn As Integer
  Dim IdxNumOfRecs As Integer
  Dim IdxLHandle As Integer
  Dim x As Integer
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim SocTip$, AlocTip$, AdvEicP$, DepCare$, NQP$
  Dim FEDWAGE#, FEDTAXWH#, SOCWAGE#, SOCTAXWH#
  Dim MedWages#, MEDTAXWH#, SocTips#, ALLOCTIP#
  Dim AdvEIC#, DEPNDCAR#, NQPLAN#, BENFBOX1#
  Dim BOX13AMT#, BOX13AM1#, STATAXWH#, STAWAGE#
  Dim BOX13AM2#, BOX13AM3#
  Dim Box13Amt1$, Box14Amt1$
  Dim StateWage$, StateTax$
  Dim dlm$
  Dim StateID As String
  
  On Error GoTo ErrorHandler
  
  dlm$ = "~"
  
  If QPTrim$(frm.fptxtStartConNum.Text = "") Then
    MsgBox "Please make an entry in the Start Control Number field"
    frm.fptxtStartConNum.SetFocus
    Exit Sub
  End If
  CtrlNumb& = CLng(frm.fptxtStartConNum.Text)
  Image1$ = "######.##"
  Image2$ = "#####.##"
  Image$ = "######"
  OpenUnitFile UHandle
  Get UHandle, , Unit(1)
  Close UHandle

  StateID = ""
  If Unit(1).UFSTATE = "VA" Then
   StateID = Mid(Unit(1).UFSTAID, 1, 2)
   StateID = StateID + QPTrim$(Unit(1).UFFEDID)
   StateID = StateID + Mid(QPTrim$(Unit(1).UFSTAID), 3, Len(QPTrim$(Unit(1).UFSTAID)))
  End If

  RptName$ = "PRRPTS\W2FORMS.RPT"
  RptTitle$ = "W-2 Forms Printing"

'  PrnCnt = 0
'  MaxPrn = 41
  OpenEmpIdxLNameFile IdxLHandle
  IdxNumOfRecs = LOF(IdxLHandle) \ 2
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For x = 1 To IdxNumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  Close IdxLHandle

  frmW2ShowPctComp.Label1 = "Printing W2 Forms"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  For cnt = 1 To IdxNumOfRecs
    Get WHandle, CLng(IdxBuff(cnt).DataRecNum), W2InfoRec(1)
    Get EHandle, CLng(IdxBuff(cnt).DataRecNum), Emp2Rec(1)
'    If Emp2Rec(1).EmpPin = "173" Then Stop
    GoSub PrintW2Form
  frmW2ShowPctComp.ShowPctComp cnt, IdxNumOfRecs
  If frmW2ShowPctComp.Out = True Then
    Close
    frmW2ShowPctComp.Out = False
    frm.cmdEscape.Enabled = True
    frm.cmdProcess.Enabled = True
    EnableCloseButton frm.hwnd, True
    Unload frmW2ShowPctComp
  End If
  Next

  Close RHandle
  Close WHandle
  Close EHandle
  Close
  
  arW2WholeForm.Show
'  arW2PrintForms.Show
  EnableCloseButton frm.hwnd, True
Exit Sub


PrintW2Form:
  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo DontPrintEm
          End If
        End If
      End If
    End If
  End If

  DidOne = DidOne + 1

  LSet PEMPCity(1) = Emp2Rec(1).EmpCity
  LSet PEmpSSN(1) = Left$(Emp2Rec(1).EmpSSN, 3) + "-" + Mid$(Emp2Rec(1).EmpSSN, 4, 2) + "-" + Right$(QPTrim$(Emp2Rec(1).EmpSSN), 4)
'start of w2 forms printing
  '                    0                      1                          2                            3
  Print #RHandle, CtrlNumb&; dlm; QPTrim$(Unit(1).UFFEDID); dlm; W2InfoRec(1).FEDWAGE; dlm; W2InfoRec(1).FEDTAXWH; dlm;
  '                           4                            5                            6
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; W2InfoRec(1).SOCWAGE; dlm; W2InfoRec(1).SOCTAXWH; dlm;
  '                              7                        8                               9                            10
  Print #RHandle, QPTrim$(Unit(1).UFADDR1); dlm; QPTrim$(Unit(1).UFADDR2); dlm; W2InfoRec(1).MedWages; dlm; W2InfoRec(1).MEDTAXWH; dlm;
  '                            11                                   12                             13
  Print #RHandle, Left$(QPTrim$(Unit(1).UFCITY), 20); dlm; QPTrim$(Unit(1).UFSTATE); dlm; QPTrim$(Unit(1).UFZIP); dlm; '11
  
  SocTip$ = W2InfoRec(1).SocTips
  AlocTip$ = W2InfoRec(1).ALLOCTIP
  '                  14            15
  Print #RHandle, SocTip$; dlm; AlocTip$; dlm; '12

  AdvEicP$ = W2InfoRec(1).AdvEIC

  DepCare$ = W2InfoRec(1).DEPNDCAR

  '                  16              17             18
  Print #RHandle, PEmpSSN(1); dlm; AdvEicP$; dlm; DepCare$; dlm; '14
  
  NQP$ = W2InfoRec(1).NQPLAN
  
  Box13Amt1$ = W2InfoRec(1).BOX13AMT   '12a

  '                            19                               20                       21
  Print #RHandle, QPTrim$(Emp2Rec(1).EmpFName); dlm; QPTrim$(Emp2Rec(1).EmpLName); dlm; NQP$; dlm;
  
  '                               22                       23                       24
  Print #RHandle, QPTrim$(W2InfoRec(1).BOX13TXt); dlm; Box13Amt1$; dlm; QPTrim$(Emp2Rec(1).EmpAddr1); dlm; '16

  Box13Amt1$ = W2InfoRec(1).BOX13AM1

  '                        25                                      26                                    27                                28
  Print #RHandle, QPTrim$(Emp2Rec(1).EMPADDR2); dlm; QPTrim$(W2InfoRec(1).BOX15A); dlm; QPTrim$(W2InfoRec(1).BOX15c); dlm; QPTrim$(W2InfoRec(1).BOX15G); dlm;
  '                            29                           30             31                         32                              33
  Print #RHandle, QPTrim$(W2InfoRec(1).BOX13TX1); dlm; Box13Amt1$; dlm; PEMPCity(1); dlm; QPTrim$(Emp2Rec(1).EmpState); dlm; QPTrim$(Emp2Rec(1).EmpZip); dlm; '18

  '8/25/04 insert 12c and 12d

  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TXT) ' + Space$(5)        'line 14 a
  Box14Amt1$ = W2InfoRec(1).BOX14AMT
  '                  34                35
  Print #RHandle, QPTrim$(BTxt14(1)); dlm; Box14Amt1$; dlm; '20


  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TX1) ' + Space$(5)        'line 14 b
  Box14Amt1$ = W2InfoRec(1).BOX14AM1
  '                  36               37
  Print #RHandle, QPTrim$(BTxt14(1)); dlm; Box14Amt1$; dlm; '21

  StateWage$ = W2InfoRec(1).STAWAGE
  StateTax$ = W2InfoRec(1).STATAXWH
  If Unit(1).UFSTATE = "VA" Then
    '                       38                 39              40              41
    Print #RHandle, W2InfoRec(1).State; dlm; StateID; dlm; StateWage$; dlm; StateTax$; dlm; '25
  Else
    '                       38                        39                        40              41
    Print #RHandle, W2InfoRec(1).State; dlm; QPTrim$(Unit(1).UFSTAID); dlm; StateWage$; dlm; StateTax$; dlm; '25
  End If
  '                       42                                   43                                      44                          45
  Print #RHandle, QPTrim$(W2InfoRec(1).BOX13TX2); dlm; W2InfoRec(1).BOX13AM2; dlm; QPTrim$(W2InfoRec(1).BOX13TX3); dlm; W2InfoRec(1).BOX13AM3
  PrnCnt = PrnCnt + 1
  CtrlNumb& = CtrlNumb& + 1

DontPrintEm:
  Return

SumW2SubTotal:
  FEDWAGE# = OldRound#(FEDWAGE# + W2InfoRec(1).FEDWAGE)
  FEDTAXWH# = OldRound#(FEDTAXWH# + W2InfoRec(1).FEDTAXWH)
  SOCWAGE# = OldRound#(SOCWAGE# + W2InfoRec(1).SOCWAGE)
  SOCTAXWH# = OldRound#(SOCTAXWH# + W2InfoRec(1).SOCTAXWH)
  MedWages# = OldRound#(MedWages# + W2InfoRec(1).MedWages)
  MEDTAXWH# = OldRound#(MEDTAXWH# + W2InfoRec(1).MEDTAXWH)
  SocTips# = OldRound#(SocTips# + W2InfoRec(1).SocTips)
  ALLOCTIP# = OldRound#(ALLOCTIP# + W2InfoRec(1).ALLOCTIP)
  AdvEIC# = OldRound#(AdvEIC# + W2InfoRec(1).AdvEIC)
  DEPNDCAR# = OldRound#(DEPNDCAR# + W2InfoRec(1).DEPNDCAR)
  NQPLAN# = OldRound#(NQPLAN# + W2InfoRec(1).NQPLAN)
  '11
  BENFBOX1# = OldRound#(BENFBOX1# + W2InfoRec(1).BENFBOX1)
  BOX13AMT# = OldRound#(BOX13AMT# + W2InfoRec(1).BOX13AMT)
  BOX13AM1# = OldRound#(BOX13AM1# + W2InfoRec(1).BOX13AM1)
  BOX13AM2# = OldRound#(BOX13AM2# + W2InfoRec(1).BOX13AM2)
  BOX13AM3# = OldRound#(BOX13AM3# + W2InfoRec(1).BOX13AM3)
  STATAXWH# = OldRound#(STATAXWH# + W2InfoRec(1).STATAXWH)
  STAWAGE# = OldRound#(STAWAGE# + W2InfoRec(1).STAWAGE)
Return

ErrorHandler:
   Unload frmW2ShowPctComp
   If Err.Number = 6 Then
     frmW2Message.Label1.Caption = "Error: some of the currency values are causing overflow problems. This is probably because extraction has not taken place. Please try extracting the latest data before continuing."
     frmW2Message.Label1.Top = 650
     frmW2Message.Show vbModal
     MainLog ("Error: User warned that during the printing for Form W2 (W2PrintGForms) some of the values are causing overflow issues and that extraction needed to be processed.")
     Close
     Exit Sub
   End If
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "W2Common", "W2PrintGForms", Erl)
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

End Sub

Sub W2ReprintTForms(frm As Form, StartNum As Long, EndNum As Long)
  ReDim PEMPCity(1) As String * 20
  ReDim PEmpSSN(1) As String * 15
  ReDim BTxt14(1) As String * 5
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  Dim W2IdxRec As W2ReprintIdxType
  Dim IdxRHandle As Integer
  Dim NumOfIdxRRecs As Integer
  Dim Image1$, Image2$, Image$
  Dim UHandle As Integer
  Dim CtrNumbStart&
  Dim CtrNumbEnd&
  Dim DidOne As Integer
  Dim RptName$, RptTitle$
  Dim PrnCnt As Integer
  Dim MaxPrn As Integer
  Dim IdxNumOfRecs As Integer
  Dim IdxLHandle As Integer
  Dim x As Integer
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim SocTip$, AlocTip$, AdvEicP$, DepCare$, NQP$
  Dim FEDWAGE#, FEDTAXWH#, SOCWAGE#, SOCTAXWH#
  Dim MedWages#, MEDTAXWH#, SocTips#, ALLOCTIP#
  Dim AdvEIC#, DEPNDCAR#, NQPLAN#, BENFBOX1#
  Dim BOX13AMT#, BOX13AM1#, STATAXWH#, STAWAGE#
  Dim BOX13AM2#, BOX13AM3#
  Dim Box13Amt1$, Box14Amt1$
  Dim StateWage$, StateTax$
  Dim StateID As String
  
  If QPTrim$(frm.fptxtStartConNum.Text = "") Then
    MsgBox "Please make an entry in the Starting Control Number field"
    frm.fptxtStartConNum.SetFocus
    Exit Sub
  End If
  If QPTrim$(frm.fptxtEndConNum.Text = "") Then
    MsgBox "Please make an entry in the Ending Control Number field"
    frm.fptxtEndConNum.SetFocus
    Exit Sub
  End If
  CtrNumbStart& = CLng(frm.fptxtStartConNum.Text)
  CtrNumbEnd& = CLng(frm.fptxtEndConNum.Text)
  Image1$ = "######.##"
  Image2$ = "#####.##"
  Image$ = "######"
  OpenUnitFile UHandle
  Get UHandle, , Unit(1)
  Close UHandle
  StateID = ""
  If Unit(1).UFSTATE = "VA" Then
   StateID = Mid(Unit(1).UFSTAID, 1, 2)
   StateID = StateID + QPTrim$(Unit(1).UFFEDID)
   StateID = StateID + Mid(QPTrim$(Unit(1).UFSTAID), 3, Len(QPTrim$(Unit(1).UFSTAID)))
  End If

  RptName$ = W2RePrintFile
  RptTitle$ = "W-2 Forms Reprinting"

  PrnCnt = 0
  MaxPrn = 41
  OpenEmpIdxLNameFile IdxLHandle
  IdxNumOfRecs = LOF(IdxLHandle) \ 2
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For x = 1 To IdxNumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  Close IdxLHandle

  frmW2ShowPctComp.Label1 = "Printing W2 Forms"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  OpenW2ReprintIdx IdxRHandle
  NumOfIdxRRecs = LOF(IdxRHandle) / Len(W2IdxRec)
  For cnt = 1 To NumOfIdxRRecs
    Get IdxRHandle, cnt, W2IdxRec
    If W2IdxRec.CONTNUM < StartNum Or W2IdxRec.CONTNUM > EndNum Then GoTo SkipThisOne
    Get WHandle, W2IdxRec.RECNO, W2InfoRec(1)
    Get EHandle, W2IdxRec.RECNO, Emp2Rec(1)
    GoSub PrintW2Form
SkipThisOne:
  
  frmW2ShowPctComp.ShowPctComp cnt, NumOfIdxRRecs
  If frmW2ShowPctComp.Out = True Then
    Close
    frmW2ShowPctComp.Out = False
    frm.cmdEscape.Enabled = True
    frm.cmdProcess.Enabled = True
    EnableCloseButton frm.hwnd, True
    Unload frmW2ShowPctComp
  End If
  Next
  
  frm.cmdEscape.Enabled = True
  frm.cmdProcess.Enabled = True
  EnableCloseButton frm.hwnd, True
  Unload frmW2ShowPctComp
  Close IdxRHandle
  Close RHandle
  Close WHandle
  Close EHandle

  W2ViewPrint RptName$, RptTitle$, True, , False
Exit Sub


PrintW2Form:
  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo DontPrintEm
          End If
        End If
      End If
    End If
  End If

  DidOne = DidOne + 1

  LSet PEMPCity(1) = Emp2Rec(1).EmpCity
  LSet PEmpSSN(1) = Left$(Emp2Rec(1).EmpSSN, 3) + "-" + Mid$(Emp2Rec(1).EmpSSN, 4, 2) + "-" + Right$(QPTrim$(Emp2Rec(1).EmpSSN), 4)
'start of w2 forms printing
  Print #RHandle, "!" '1
  Print #RHandle,     '2
  Print #RHandle,     '3 'added 10/13/04 for 2004 forms
  Print #RHandle,     '4
'  Print #RHandle, Tab(6); Using(Image$, CtrNumbStart&) '5
  Print #RHandle, Tab(23); PEmpSSN(1) '5
  Print #RHandle,     '6
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFFEDID); Tab(51); Using(Image1$, W2InfoRec(1).FEDWAGE); Tab(68); Using(Image1$, W2InfoRec(1).FEDTAXWH) '7
  Print #RHandle, '8
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFEMPR); Tab(51); Using(Image1$, W2InfoRec(1).SOCWAGE); Tab(68); Using(Image1$, W2InfoRec(1).SOCTAXWH) '9
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFADDR1) '10
  Print #RHandle, Tab(5); QPTrim$(Unit(1).UFADDR2); Tab(51); Using(Image1$, W2InfoRec(1).MedWages); Tab(68); Using(Image1$, W2InfoRec(1).MEDTAXWH) '11
  Print #RHandle, Tab(5); Left$(QPTrim$(Unit(1).UFCITY), 15) + " " + QPTrim$(Unit(1).UFSTATE) + " " + QPTrim$(Unit(1).UFZIP) '12
  
  If W2InfoRec(1).SocTips > 0 Then
    SocTip$ = Using(Image1$, W2InfoRec(1).SocTips)
  Else
    SocTip$ = Space$(Len(Image1$))
  End If
  If W2InfoRec(1).ALLOCTIP > 0 Then
    AlocTip$ = Using(Image1$, W2InfoRec(1).ALLOCTIP)
  Else
    AlocTip$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(51); SocTip$; Tab(68); AlocTip$ '13

  Print #RHandle,     ' line 14

  If W2InfoRec(1).AdvEIC > 0 Then
    AdvEicP$ = Using(Image1$, W2InfoRec(1).AdvEIC)
  Else
    AdvEicP$ = Space$(Len(Image1$))
  End If

  If W2InfoRec(1).DEPNDCAR > 0 Then
    DepCare$ = Using(Image1$, W2InfoRec(1).DEPNDCAR)
  Else
    DepCare$ = Space$(Len(Image1$))
  End If

'  Print #RHandle, Tab(5); PEmpSSN(1); Tab(51); AdvEicP$; Tab(68); DepCare$ '15
  Print #RHandle, Tab(15); Using(Image$, CtrNumbStart&); Tab(51); AdvEicP$; Tab(68); DepCare$ '15
  Print #RHandle, '16

  If W2InfoRec(1).NQPLAN > 0 Then
    NQP$ = Using(Image1$, W2InfoRec(1).NQPLAN)
  Else
    NQP$ = Space$(Len(Image1$))
  End If
  
  If W2InfoRec(1).BOX13AMT > 0 Then    '12a
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AMT)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(5); QPTrim$(Emp2Rec(1).EmpFName); Tab(28); QPTrim$(Emp2Rec(1).EmpLName); Tab(51); NQP$;
'  Print #RHandle, Tab(66); Left$(QPTrim$(W2InfoRec(1).BOX13TXt) + Space$(1), 1); "   "; Box13Amt1$ '16 changed for 2004 forms
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TXt) + Space$(1), 1); "   "; Box13Amt1$ '17
  
  Print #RHandle, 'Tab(5); QPTrim$(Emp2Rec(1).EmpAddr1) '18

  If W2InfoRec(1).BOX13AM1 > 0 Then
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM1)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If

  Print #RHandle, Tab(5); QPTrim$(Emp2Rec(1).EmpAddr1);
'  Print #RHandle, Tab(49); QPTrim$(W2InfoRec(1).BOX15A); Tab(54); QPTrim$(W2InfoRec(1).BOX15c); Tab(59); QPTrim$(W2InfoRec(1).BOX15G); changed for 2004 forms
  Print #RHandle, Tab(48); QPTrim$(W2InfoRec(1).BOX15A); Tab(53); QPTrim$(W2InfoRec(1).BOX15c); Tab(58); QPTrim$(W2InfoRec(1).BOX15G);
'  Print #RHandle, Tab(66); Left$(QPTrim$(W2InfoRec(1).BOX13TX1) + Space$(1), 1); "   "; Box13Amt1$ '18 changed for 2004 forms
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TX1) + Space$(1), 1); "   "; Box13Amt1$ '19

'  '08/25/04 insert 12c and 12d
  
  Print #RHandle, QPTrim$(Emp2Rec(1).EMPADDR2) '20

  If W2InfoRec(1).BOX13AM2 > 0 Then 'added fall 04
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM2)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If
  
  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TXT) + Space$(5)        'line 14 a
  If W2InfoRec(1).BOX14AMT > 0 Then
    Box14Amt1$ = Using(Image1$, W2InfoRec(1).BOX14AMT)
  Else
    Box14Amt1$ = Space$(Len(Image1$))
  End If
  Print #RHandle, Tab(5); PEMPCity(1) + " "; QPTrim$(Emp2Rec(1).EmpState); " "; QPTrim$(Emp2Rec(1).EmpZip); Tab(47); BTxt14(1); Box14Amt1$; '21
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TX2) + Space$(1), 1); "   "; Box13Amt1$ '21'added fall 04

  If W2InfoRec(1).BOX13AM3 > 0 Then 'added fall 04
    Box13Amt1$ = Using(Image1$, W2InfoRec(1).BOX13AM3)
  Else
    Box13Amt1$ = Space$(Len(Image1$))
  End If
  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TX1) + Space$(5)        'line 14 b
  If W2InfoRec(1).BOX14AM1 > 0 Then
    Box14Amt1$ = Using(Image1$, W2InfoRec(1).BOX14AM1)
  Else
    Box14Amt1$ = Space$(Len(Image1$))
  End If
  Print #RHandle, Tab(47); BTxt14(1); Box14Amt1$ '22
'  Print #RHandle, '22
  Print #RHandle, Tab(65); Left$(QPTrim$(W2InfoRec(1).BOX13TX3) + Space$(1), 1); "   "; Box13Amt1$ '23'added fall 04
  Print #RHandle, '24

  If W2InfoRec(1).STAWAGE > 0 Then
    StateWage$ = Using(Image1$, W2InfoRec(1).STAWAGE)
  Else
    StateWage$ = Space$(Len(Image1$))
  End If

  If W2InfoRec(1).STATAXWH > 0 Then
    StateTax$ = Using(Image1$, W2InfoRec(1).STATAXWH)
  Else
    StateTax$ = Space$(Len(Image1$))
  End If

  Print #RHandle, '25
'  Print #RHandle, Tab(3); W2InfoRec(1).State; Tab(9); QPTrim$(Unit(1).UFSTAID); Tab(26); StateWage$; Tab(38); StateTax$ '25 changed for 2004 forms
  If Unit(1).UFSTATE = "VA" Then
    Print #RHandle, Tab(5); W2InfoRec(1).State; Tab(11); StateID; Tab(28); StateWage$; Tab(39); StateTax$ '26
  Else
    Print #RHandle, Tab(5); W2InfoRec(1).State; Tab(11); QPTrim$(Unit(1).UFSTAID); Tab(28); StateWage$; Tab(39); StateTax$ '26
  End If
  Print #RHandle, '27
  Print #RHandle, '28
  Print #RHandle, '29
  Print #RHandle, '30
  Print #RHandle, '31
  Print #RHandle, '32
'  Print #RHandle, '32 'omitted for 2004 forms
  Print #RHandle, "!" '33

  PrnCnt = PrnCnt + 1
  CtrNumbStart& = CtrNumbStart& + 1

DontPrintEm:
  Return

SumW2SubTotal:
  FEDWAGE# = OldRound#(FEDWAGE# + W2InfoRec(1).FEDWAGE)
  FEDTAXWH# = OldRound#(FEDTAXWH# + W2InfoRec(1).FEDTAXWH)
  SOCWAGE# = OldRound#(SOCWAGE# + W2InfoRec(1).SOCWAGE)
  SOCTAXWH# = OldRound#(SOCTAXWH# + W2InfoRec(1).SOCTAXWH)
  MedWages# = OldRound#(MedWages# + W2InfoRec(1).MedWages)
  MEDTAXWH# = OldRound#(MEDTAXWH# + W2InfoRec(1).MEDTAXWH)
  SocTips# = OldRound#(SocTips# + W2InfoRec(1).SocTips)
  ALLOCTIP# = OldRound#(ALLOCTIP# + W2InfoRec(1).ALLOCTIP)
  AdvEIC# = OldRound#(AdvEIC# + W2InfoRec(1).AdvEIC)
  DEPNDCAR# = OldRound#(DEPNDCAR# + W2InfoRec(1).DEPNDCAR)
  NQPLAN# = OldRound#(NQPLAN# + W2InfoRec(1).NQPLAN)
  '11
  BENFBOX1# = OldRound#(BENFBOX1# + W2InfoRec(1).BENFBOX1)
  BOX13AMT# = OldRound#(BOX13AMT# + W2InfoRec(1).BOX13AMT)
  BOX13AM1# = OldRound#(BOX13AM1# + W2InfoRec(1).BOX13AM1)
  BOX13AM2# = OldRound#(BOX13AM2# + W2InfoRec(1).BOX13AM2)
  BOX13AM3# = OldRound#(BOX13AM3# + W2InfoRec(1).BOX13AM3)
  STATAXWH# = OldRound#(STATAXWH# + W2InfoRec(1).STATAXWH)
  STAWAGE# = OldRound#(STAWAGE# + W2InfoRec(1).STAWAGE)
Return

End Sub

Sub W2PrintG4Forms(frm As Form, FormType$)
  ReDim PEMPCity(1) As String * 20
  ReDim PEmpSSN(1) As String * 15
  ReDim BTxt14(1) As String * 5
  ReDim Unit(1) As UnitFileRecType
  ReDim W2InfoRec(1) As W2FormType
  Dim Image1$, Image2$, Image$
  Dim UHandle As Integer
  Dim CtrlNumb&
  Dim DidOne As Integer
  Dim RptName$, RptTitle$
  Dim PrnCnt As Integer
  Dim MaxPrn As Integer
  Dim IdxNumOfRecs As Integer
  Dim IdxLHandle As Integer
  Dim x As Integer
  Dim RHandle As Integer
  Dim WHandle As Integer
  Dim EHandle As Integer
  Dim cnt As Integer
  Dim SocTip$, AlocTip$, AdvEicP$, DepCare$, NQP$
  Dim FEDWAGE#, FEDTAXWH#, SOCWAGE#, SOCTAXWH#
  Dim MedWages#, MEDTAXWH#, SocTips#, ALLOCTIP#
  Dim AdvEIC#, DEPNDCAR#, NQPLAN#, BENFBOX1#
  Dim BOX13AMT#, BOX13AM1#, STATAXWH#, STAWAGE#
  Dim BOX13AM2#, BOX13AM3#
  Dim Box13Amt1$, Box14Amt1$
  Dim StateWage$, StateTax$
  Dim dlm$
  Dim StateID As String
  
  On Error GoTo ErrorHandler
  
  dlm$ = "~"
  
  If QPTrim$(frm.fptxtStartConNum.Text = "") Then
    MsgBox "Please make an entry in the Start Control Number field"
    frm.fptxtStartConNum.SetFocus
    Exit Sub
  End If
  CtrlNumb& = CLng(frm.fptxtStartConNum.Text)
  Image1$ = "######.##"
  Image2$ = "#####.##"
  Image$ = "######"
  OpenUnitFile UHandle
  Get UHandle, , Unit(1)
  Close UHandle
  StateID = ""
  If Unit(1).UFSTATE = "VA" Then
   StateID = Mid(Unit(1).UFSTAID, 1, 2)
   StateID = StateID + QPTrim$(Unit(1).UFFEDID)
   StateID = StateID + Mid(QPTrim$(Unit(1).UFSTAID), 3, Len(QPTrim$(Unit(1).UFSTAID)))
  End If

  RptName$ = "PRRPTS\W2FRMS4.RPT"
  RptTitle$ = "W-2 Forms Printing"

  OpenEmpIdxLNameFile IdxLHandle
  IdxNumOfRecs = LOF(IdxLHandle) \ 2
  ReDim IdxBuff(1 To IdxNumOfRecs) As EmployeeIndexType         'load index file
  For x = 1 To IdxNumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  Close IdxLHandle

  frmW2ShowPctComp.Label1 = "Printing W2 Forms"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle

  OpenW2Info WHandle
  OpenEmpData2File EHandle
  For cnt = 1 To IdxNumOfRecs
    Get WHandle, CLng(IdxBuff(cnt).DataRecNum), W2InfoRec(1)
    Get EHandle, CLng(IdxBuff(cnt).DataRecNum), Emp2Rec(1)
    GoSub PrintW2Form
  frmW2ShowPctComp.ShowPctComp cnt, IdxNumOfRecs
  If frmW2ShowPctComp.Out = True Then
    Close
    frmW2ShowPctComp.Out = False
    frm.cmdEscape.Enabled = True
    frm.cmdProcess.Enabled = True
    EnableCloseButton frm.hwnd, True
    Unload frmW2ShowPctComp
  End If
  Next

  Close RHandle
  Close WHandle
  Close EHandle
  Close
  
  If FormType = "Corners" Then
    arW2PrintForms4.Show
  ElseIf FormType = "Horizontal" Then
    arW2PrintFormsH4.Show
  End If
  EnableCloseButton frm.hwnd, True
Exit Sub


PrintW2Form:
  If W2InfoRec(1).FEDWAGE = 0 And W2InfoRec(1).FEDTAXWH = 0 And W2InfoRec(1).SOCWAGE = 0 Then
    If W2InfoRec(1).SOCTAXWH = 0 And W2InfoRec(1).MedWages = 0 And W2InfoRec(1).MEDTAXWH = 0 Then
      If W2InfoRec(1).SocTips = 0 And W2InfoRec(1).ALLOCTIP = 0 And W2InfoRec(1).AdvEIC = 0 Then
        If W2InfoRec(1).DEPNDCAR = 0 And W2InfoRec(1).NQPLAN = 0 And W2InfoRec(1).BOX13AMT = 0 Then
          If W2InfoRec(1).BOX13AM1 = 0 And W2InfoRec(1).BOX13AM2 = 0 And W2InfoRec(1).BOX13AM3 = 0 Then
            GoTo DontPrintEm
          End If
        End If
      End If
    End If
  End If

  DidOne = DidOne + 1


  LSet PEMPCity(1) = Emp2Rec(1).EmpCity
  LSet PEmpSSN(1) = Left$(Emp2Rec(1).EmpSSN, 3) + "-" + Mid$(Emp2Rec(1).EmpSSN, 4, 2) + "-" + Right$(QPTrim$(Emp2Rec(1).EmpSSN), 4)
  '                    0                      1                          2                            3
  Print #RHandle, CtrlNumb&; dlm; QPTrim$(Unit(1).UFFEDID); dlm; W2InfoRec(1).FEDWAGE; dlm; W2InfoRec(1).FEDTAXWH; dlm;
  
  '                           4                            5                            6
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; W2InfoRec(1).SOCWAGE; dlm; W2InfoRec(1).SOCTAXWH; dlm;
  '                              7                        8                               9                            10
  Print #RHandle, QPTrim$(Unit(1).UFADDR1); dlm; QPTrim$(Unit(1).UFADDR2); dlm; W2InfoRec(1).MedWages; dlm; W2InfoRec(1).MEDTAXWH; dlm;
  '                            11                                   12                             13
  Print #RHandle, Left$(QPTrim$(Unit(1).UFCITY), 20); dlm; QPTrim$(Unit(1).UFSTATE); dlm; QPTrim$(Unit(1).UFZIP); dlm; '11
  
  
  SocTip$ = W2InfoRec(1).SocTips
  AlocTip$ = W2InfoRec(1).ALLOCTIP
  '                  14            15
  Print #RHandle, SocTip$; dlm; AlocTip$; dlm; '12
  AdvEicP$ = W2InfoRec(1).AdvEIC
  DepCare$ = W2InfoRec(1).DEPNDCAR

  '                  16              17             18
  Print #RHandle, PEmpSSN(1); dlm; AdvEicP$; dlm; DepCare$; dlm; '14
  NQP$ = W2InfoRec(1).NQPLAN
  Box13Amt1$ = W2InfoRec(1).BOX13AMT
  '                            19                               20                       21
  Print #RHandle, QPTrim$(Emp2Rec(1).EmpFName); dlm; QPTrim$(Emp2Rec(1).EmpLName); dlm; NQP$; dlm;
  '                               22                       23                       24
  Print #RHandle, QPTrim$(W2InfoRec(1).BOX13TXt); dlm; Box13Amt1$; dlm; QPTrim$(Emp2Rec(1).EmpAddr1); dlm; '16
  Box13Amt1$ = W2InfoRec(1).BOX13AM1
  '                        25                                      26                                    27                                28
  Print #RHandle, QPTrim$(Emp2Rec(1).EMPADDR2); dlm; QPTrim$(W2InfoRec(1).BOX15A); dlm; QPTrim$(W2InfoRec(1).BOX15c); dlm; QPTrim$(W2InfoRec(1).BOX15G); dlm;
  '                            29                           30             31                         32                              33
  Print #RHandle, QPTrim$(W2InfoRec(1).BOX13TX1); dlm; Box13Amt1$; dlm; PEMPCity(1); dlm; QPTrim$(Emp2Rec(1).EmpState); dlm; QPTrim$(Emp2Rec(1).EmpZip); dlm; '18

  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TXT) ' + Space$(5)        'line 14 a
  Box14Amt1$ = W2InfoRec(1).BOX14AMT
  '                  34                35
  Print #RHandle, QPTrim$(BTxt14(1)); dlm; Box14Amt1$; dlm; '20

  BTxt14(1) = QPTrim$(W2InfoRec(1).BOX14TX1) ' + Space$(5)        'line 14 b
  Box14Amt1$ = W2InfoRec(1).BOX14AM1
  '                  36               37
  Print #RHandle, QPTrim$(BTxt14(1)); dlm; Box14Amt1$; dlm; '21
  StateWage$ = W2InfoRec(1).STAWAGE
  StateTax$ = W2InfoRec(1).STATAXWH
  If Unit(1).UFSTATE = "VA" Then
    '                       38                   39             40              41
    Print #RHandle, W2InfoRec(1).State; dlm; StateID; dlm; StateWage$; dlm; StateTax$; dlm; '25
  Else
    '                       38                        39                        40              41
    Print #RHandle, W2InfoRec(1).State; dlm; QPTrim$(Unit(1).UFSTAID); dlm; StateWage$; dlm; StateTax$; dlm; '25
  End If
'  '8/25/04 insert 12c and 12d
  '                   42                            43
  Print #RHandle, W2InfoRec(1).BOX13AM2; dlm; QPTrim$(W2InfoRec(1).BOX13TX2); dlm;
  '                   44                            45
  Print #RHandle, W2InfoRec(1).BOX13AM3; dlm; QPTrim$(W2InfoRec(1).BOX13TX3)
  
  PrnCnt = PrnCnt + 1
  CtrlNumb& = CtrlNumb& + 1

DontPrintEm:
  Return

SumW2SubTotal:
  FEDWAGE# = OldRound#(FEDWAGE# + W2InfoRec(1).FEDWAGE)
  FEDTAXWH# = OldRound#(FEDTAXWH# + W2InfoRec(1).FEDTAXWH)
  SOCWAGE# = OldRound#(SOCWAGE# + W2InfoRec(1).SOCWAGE)
  SOCTAXWH# = OldRound#(SOCTAXWH# + W2InfoRec(1).SOCTAXWH)
  MedWages# = OldRound#(MedWages# + W2InfoRec(1).MedWages)
  MEDTAXWH# = OldRound#(MEDTAXWH# + W2InfoRec(1).MEDTAXWH)
  SocTips# = OldRound#(SocTips# + W2InfoRec(1).SocTips)
  ALLOCTIP# = OldRound#(ALLOCTIP# + W2InfoRec(1).ALLOCTIP)
  AdvEIC# = OldRound#(AdvEIC# + W2InfoRec(1).AdvEIC)
  DEPNDCAR# = OldRound#(DEPNDCAR# + W2InfoRec(1).DEPNDCAR)
  NQPLAN# = OldRound#(NQPLAN# + W2InfoRec(1).NQPLAN)
  '11
  BENFBOX1# = OldRound#(BENFBOX1# + W2InfoRec(1).BENFBOX1)
  BOX13AMT# = OldRound#(BOX13AMT# + W2InfoRec(1).BOX13AMT)
  BOX13AM1# = OldRound#(BOX13AM1# + W2InfoRec(1).BOX13AM1)
  BOX13AM2# = OldRound#(BOX13AM2# + W2InfoRec(1).BOX13AM2)
  BOX13AM3# = OldRound#(BOX13AM3# + W2InfoRec(1).BOX13AM3)
  STATAXWH# = OldRound#(STATAXWH# + W2InfoRec(1).STATAXWH)
  STAWAGE# = OldRound#(STAWAGE# + W2InfoRec(1).STAWAGE)

Return

ErrorHandler:
   Unload frmW2ShowPctComp
   If Err.Number = 6 Then
     frmW2Message.Label1.Caption = "Error: some of the currency values are causing overflow problems. This is probably because extraction has not taken place. Please try extracting the latest data before continuing."
     frmW2Message.Label1.Top = 650
     frmW2Message.Show vbModal
     MainLog ("Error: User warned that during the printing for Form W2 (W2PrintG4Forms) some of the values are causing overflow issues and that extraction needed to be processed.")
     Close
     Exit Sub
   End If
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "W2Common", "W2PrintG4Forms", Erl)
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

End Sub

Public Sub ConvertNames(ByRef FName$, ByRef MName$, ByRef LName$, ByRef Suffix$)
  Dim x As Integer
  Dim FTmpName As String * 15
  Dim FStop As Integer
  Dim MTmpName As String * 15
  Dim LTmpName As String * 20
  Dim TmpSfx As String * 4
  Dim LStop As Integer
  Dim thischar$
  Dim FLen As Integer
  Dim LLen As Integer
  
  LSet FTmpName = QPTrim$(FName)
  LSet LTmpName = QPTrim(LName)
  FLen = Len(QPTrim$(FName))
  For x = 1 To FLen
    thischar = Mid$(FTmpName, x, 1)
    If thischar = " " Then
      FStop = x
      Exit For
    End If
  Next x
  
  If FStop > 0 Then
    MTmpName = Mid(FTmpName, FStop + 1, FLen)
    LSet MTmpName = MTmpName
    FTmpName = Mid(FName, 1, FStop - 1)
    LSet FTmpName = FTmpName
  Else
    MTmpName = "               "
    LSet FTmpName = FTmpName
  End If
  
  LLen = Len(QPTrim(LName))
  For x = 1 To LLen
    thischar = Mid$(LName, x, 1)
    If thischar = " " Then
      LStop = x
      Exit For
    End If
  Next x
  
  If LStop > 0 Then
    TmpSfx = Mid(LTmpName, LStop + 1, LLen)
    LSet TmpSfx = TmpSfx
    LTmpName = Mid(LTmpName, 1, LStop - 1)
    LSet LTmpName = LTmpName
  Else
    TmpSfx = "    "
    LSet LTmpName = LTmpName
  End If
  FName = ReplaceString(FTmpName, ",", "")
  MName = ReplaceString(MTmpName, ",", "")
  LName = ReplaceString(LTmpName, ",", "")
  Suffix = ReplaceString(TmpSfx, ",", "")
End Sub

Public Sub ZeroFill(ByRef ThisNum$, ThisLen As Integer)
  Dim x As Integer
  Dim thischar$
  Dim BCnt As Integer
  Dim ThisTemp$
  
  For x = 1 To ThisLen
    thischar = Mid(ThisNum, x, 1)
    If thischar = " " Then
      BCnt = BCnt + 1
    End If
  Next x
  
  For x = 1 To BCnt
    ThisTemp = ThisTemp + "0"
  Next x
  
  ThisNum$ = ThisTemp + QPTrim$(ThisNum)
  
  
End Sub

Public Function Currency2String$(Amt As Double, ThisLen As Integer)
  Dim StrLen As Long
  Dim cnt As Integer
  Dim StrAmt As String
  Dim NewText As String
  Dim thischar$
  Dim CTChar$
  Dim TTChar$
  Dim CTLen As Integer
  Dim TTLen As Integer
  Dim BigLen As Integer
  Dim StopHere As Integer
  Dim Nextx As Integer
  Dim DifCnt As Integer
  
  If Amt > 10000000000000# Then
    MsgBox "The amount " + CStr(Amt) + " is greater than the largest amount allowed."
    Currency2String = "Error"
    Exit Function
  End If
  StrAmt = Using$("###########0.00", Amt)
  StrAmt = QPTrim$(StrAmt)
  StrLen = ThisLen
  ReDim AmtStr(1 To 15) As String
  
  For cnt = 1 To Len(StrAmt)
    thischar = Mid$(StrAmt, cnt, 1)
    If thischar = "." Then GoTo SkipDecimal
    Nextx = Nextx + 1
    AmtStr(Nextx) = thischar
SkipDecimal:
  Next
  
  For cnt = Nextx + 1 To ThisLen
    Currency2String = Currency2String + "0"
  Next cnt
  
  For cnt = 1 To Nextx
    Currency2String = Currency2String + AmtStr(cnt)
  Next cnt
  
End Function

Sub ExtractW2Info(W2Type%, frm As Form)

  ReDim sp2(1) As String * 2
  ReDim W2SetUpRec(1) As W2SetUpType
  Dim W2Handle As Integer
  Dim W2SetUpRecLen As Integer
  Dim DedRec As DedCodeRecType
  Dim DedRecCnt As Integer
  Dim DedHandle As Integer
  Dim x As Integer
  Dim W2Year$
  Dim StrDate As Integer
  Dim EndDate As Integer
  Dim UnitRecLen As Integer
  Dim W2FormRecLen As Integer
  Dim Emp2RecLen As Integer
  Dim TranRecLen As Integer
  Dim ENumOfRec As Integer
  Dim TNumOfRec As Long
  Dim RptTitle$
  Dim UnitHandle As Integer
  Dim THandle As Integer
  Dim EHandle As Integer
  Dim WHandle As Integer
  Dim cnt As Long
  Dim ECnt As Integer
  Dim TotalTransRecs As Long
  Dim TCnt As Long
  Dim CntZZ As Integer
  Dim W2RWRec As W2ElectronicSubRW
  Dim BW2RWRec As W2ElectronicSubRW
  Dim RWHandle As Integer
  Dim FedTax As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  Dim ThisLifeIns As Double
  Dim ThisDefr401k As Double
  Dim ThisDefr403b As Double
  Dim ThisDefr408k6 As Double
  Dim ThisDefr457b As Double
  Dim ThisDefr501c18D As Double
  Dim ThisNonStaStcks As Double
  Dim ThisRoth401K As Double 'added for 2006
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FedTax
  Close FedTaxHandle
  FedSSMax = FedTax.FTMSSMW
  
  OpenDedCodeFile DedHandle
  DedRecCnt = LOF(DedHandle) / Len(DedRec)
  Close DedHandle
  
  ReDim Deds(0 To 51) As Double
  
  OpenW2SetUp W2Handle
  W2SetUpRec(1).ExtrYear = frm.fptxtYear.Text
'  W2SetUpRec(1).Deds(0).CHKDED = QPTrim$(frm.fpcomboRetire.Text)
  For x = 1 To DedRecCnt + 2
    frm.vaSpreadW2.Col = 2
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).CHKDED = QPTrim$(frm.vaSpreadW2.Text)
    frm.vaSpreadW2.Col = 3
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).AMTBOX = QPTrim$(frm.vaSpreadW2.Text)
    frm.vaSpreadW2.Col = 4
    frm.vaSpreadW2.Row = x
    W2SetUpRec(1).Deds(x - 1).DedCode = QPTrim$(frm.vaSpreadW2.Text)
  Next x
  Put W2Handle, 1, W2SetUpRec(1)
  Close W2Handle

  W2Year = Val(frm.fptxtYear.Text)
  StrDate = Date2Num("01-01-" + frm.fptxtYear.Text)
  EndDate = Date2Num("12-31-" + frm.fptxtYear.Text)
  
  ReDim TranRec(1) As TransRecType
  ReDim UnitRec(1) As UnitFileRecType
  ReDim TPntr(0 To 200) As Long
  ReDim W2FormRec(1) As W2FormType
  ReDim BW2FormRec(1) As W2FormType
  
  UnitRecLen = Len(UnitRec(1))
  W2FormRecLen = Len(W2FormRec(1))
  Emp2RecLen = Len(Emp2Rec(1))
  TranRecLen = Len(TranRec(1))
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitRec(1)
  Close UnitHandle
  
  RptTitle$ = "Extracting W-2 Information"
  frmW2ShowPctComp.CmdCancel.Visible = False
  frmW2ShowPctComp.Label1 = "Extracting W-2 Information"
  frmW2ShowPctComp.Show , frm
  DoEvents
  EnableCloseButton frm.hwnd, False
  OpenTransHistFile THandle
  TNumOfRec = LOF(THandle) \ Len(TranRec(1))

  'get trans action history pins
  ReDim TPins(1 To TNumOfRec) As Integer
  For cnt = 1 To TNumOfRec
    Get THandle, cnt, TranRec(1)
    TPins(cnt) = TranRec(1).EmpPin
  Next
  
  OpenEmpData2File EHandle
  ENumOfRec = LOF(EHandle) \ Len(Emp2Rec(1))
  
  OpenW2Info WHandle
  KillFile ("PRDATA\W2ESUBRW.DAT")
  KillFile ("PRDATA\W2ESUBRT.DAT")
  KillFile ("PRDATA\W2ESUBRO.DAT")
  KillFile ("PRDATA\W2ESUBRU.DAT")
  KillFile ("PRDATA\W2ESUBRF.DAT")
  
  OpenW2ESubRW RWHandle
  
  For ECnt = 1 To ENumOfRec
    Get EHandle, ECnt, Emp2Rec(1) '12/03
    GoSub GetEmpTranRecNums
    Select Case W2Type%
    Case 1
      If Emp2Rec(1).EMPSOCX <> "Y" Then
        GoSub CheckEmp
      Else
        W2FormRec(1) = BW2FormRec(1)
        W2RWRec = BW2RWRec
      End If
    Case 2
      If Emp2Rec(1).EMPSOCX = "Y" Then
        GoSub CheckEmp
      Else
        W2FormRec(1) = BW2FormRec(1)
        W2RWRec = BW2RWRec
      End If
    End Select
    
    Put WHandle, ECnt, W2FormRec(1)
    Put RWHandle, ECnt, W2RWRec '12/03
    
    frmW2ShowPctComp.ShowPctComp ECnt, ENumOfRec
'    Get RWHandle, ECnt, W2RWRec
'    W2RWRec.WageTips = W2RWRec.WageTips
  Next
  
  Unload frmW2ShowPctComp 'added 5/28/04
  EnableCloseButton frm.hwnd, True
  
  Close EHandle
  Close WHandle
  Close THandle
  Close RWHandle
  
  GoTo ExitW2SetUp
  
GetEmpTranRecNums:

  ReDim TPntr(0 To 1600)
  TotalTransRecs = 0
  For TCnt = 1 To TNumOfRec
    If TPins(TCnt) = Emp2Rec(1).EmpPin Then
      TotalTransRecs = TotalTransRecs + 1
      TPntr(TotalTransRecs) = TCnt
    End If
    TPntr(0) = TotalTransRecs
  Next
  Return

SumW2Info:
  W2FormRec(1) = BW2FormRec(1)
  W2RWRec = BW2RWRec
  For CntZZ = 0 To 51
    Deds(CntZZ) = 0
  Next

  For cnt = 1 To TPntr(0)
    Get THandle, TPntr(cnt), TranRec(1)
    If TranRec(1).CheckDate >= StrDate And TranRec(1).CheckDate <= EndDate Then
      W2FormRec(1).FEDWAGE = OldRound#(W2FormRec(1).FEDWAGE + TranRec(1).FedGrossPay)
      W2RWRec.WageTips = Currency2String(W2FormRec(1).FEDWAGE, 11) '12/03
      
      W2FormRec(1).FEDTAXWH = OldRound#(W2FormRec(1).FEDTAXWH + TranRec(1).FedTaxAmt)
      W2RWRec.FedTax = Currency2String(W2FormRec(1).FEDTAXWH, 11)
      
      W2FormRec(1).SOCWAGE = OldRound#(W2FormRec(1).SOCWAGE + TranRec(1).SocGrossPay)
      If W2FormRec(1).SOCWAGE > FedSSMax Then
        W2FormRec(1).SOCWAGE = FedSSMax
      End If
      W2RWRec.SSWages = Currency2String(W2FormRec(1).SOCWAGE, 11) '12/03
      W2FormRec(1).SOCTAXWH = OldRound#(W2FormRec(1).SOCTAXWH + TranRec(1).SocTaxAmt)
      W2RWRec.SSTax = Currency2String(W2FormRec(1).SOCTAXWH, 11) '12/03
      
      W2FormRec(1).MedWages = OldRound#(W2FormRec(1).MedWages + TranRec(1).MedGrossPay)
      W2RWRec.MedWages = Currency2String(W2FormRec(1).MedWages, 11) '12/03
      
      W2FormRec(1).MEDTAXWH = OldRound#(W2FormRec(1).MEDTAXWH + TranRec(1).MedTaxAmt)
      W2RWRec.MedTax = Currency2String(W2FormRec(1).MEDTAXWH, 11) '12/03
      
      W2FormRec(1).AdvEIC = OldRound#(W2FormRec(1).AdvEIC + TranRec(1).EICAmt)
      W2RWRec.AdvEIC = Currency2String(W2FormRec(1).AdvEIC, 11) '12/03
      
      W2FormRec(1).BENFBOX1 = OldRound#(W2FormRec(1).BENFBOX1 + TranRec(1).TaxFring)
      
      W2FormRec(1).State = UnitRec(1).UFSTATE
      W2FormRec(1).STAWAGE = OldRound#(W2FormRec(1).STAWAGE + TranRec(1).StaGrossPay)
      W2FormRec(1).STATAXWH = OldRound#(W2FormRec(1).STATAXWH + TranRec(1).StaTaxAmt)
      Deds(0) = OldRound#(Deds(0) + TranRec(1).RetireAmt)
'      TaxFringe04 entry if needed = Deds(1) + TransRec(1).TaxFring
'      you'll need to expand the 50 to 51 everywhere in the program to
'      accomodate the tax fringe addition
      Deds(1) = OldRound#(Deds(1) + TranRec(1).TaxFring)
      For CntZZ = 2 To 51
'        Deds(CntZZ - 1) = OldRound#(Deds(CntZZ - 1) + TranRec(1).DAmt(CntZZ - 1))
        Deds(CntZZ) = OldRound#(Deds(CntZZ) + TranRec(1).DAmt(CntZZ - 1))
      Next
    End If
  Next
  '12/10/07
  W2RWRec.EmpFName = Emp2Rec(1).EmpFName
  W2RWRec.EmpMName = ""
  W2RWRec.EmpLName = Emp2Rec(1).EmpLName
  W2RWRec.EmpSuffix = ""
  Call ConvertNames(W2RWRec.EmpFName, W2RWRec.EmpMName, W2RWRec.EmpLName, W2RWRec.EmpSuffix)
  W2RWRec.EmpCity = QPTrim$(Emp2Rec(1).EmpCity) '12/03
  W2RWRec.EmpState = QPTrim$(Emp2Rec(1).EmpState) '12/03
  W2RWRec.EmpZip = QPTrim(Emp2Rec(1).EmpZip)
  W2RWRec.EmpZip = Mid(W2RWRec.EmpZip, 1, 5) '12/03
  W2RWRec.EmpZipX = QPTrim(Emp2Rec(1).EmpZip)
  W2RWRec.EmpZipX = Mid(W2RWRec.EmpZipX, 7, 4) '12/03
  W2RWRec.EmpSSN = ReplaceString(Emp2Rec(1).EmpSSN, "-", "") '12/03
  W2RWRec.EmpAdd1 = QPTrim$(Emp2Rec(1).EmpAddr1) '12/03
  W2RWRec.EmpAdd2 = QPTrim$(Emp2Rec(1).EMPADDR2) '12/03
  W2RWRec.SSTips = ""
  W2RWRec.DepCare = ""
  W2RWRec.NQPlan457 = ""
  W2RWRec.NQPNot457 = ""
  W2RWRec.StatuEmp = "0"
  W2RWRec.ThrdSckPay = "0"
  W2RWRec.ThrdSckAmt = "0"
  W2RWRec.RONum = "0"
  W2RWRec.Roth401K = "0" 'added for 2006
Return

ApplyForm:
'*******************************
  If W2FormRec(1).BENFBOX1 > 0 Then 'BENFBOX1 = Taxable Fringe
    W2FormRec(1).FEDWAGE = OldRound#(W2FormRec(1).FEDWAGE + W2FormRec(1).BENFBOX1)
    W2RWRec.WageTips = Currency2String(W2FormRec(1).FEDWAGE, 11) '12/03
    
    If Emp2Rec(1).EMPSOCX = "N" Then '7/2/2010
      W2FormRec(1).SOCWAGE = OldRound#(W2FormRec(1).SOCWAGE + W2FormRec(1).BENFBOX1)
    Else 'added this part 7/2/2010
      W2FormRec(1).SOCWAGE = W2FormRec(1).SOCWAGE
    End If
    If W2FormRec(1).SOCWAGE > FedSSMax Then
      W2FormRec(1).SOCWAGE = FedSSMax
    End If
    W2RWRec.SSWages = Currency2String(W2FormRec(1).SOCWAGE, 11) '12/03
    
    W2FormRec(1).MedWages = OldRound#(W2FormRec(1).MedWages + W2FormRec(1).BENFBOX1)
    W2RWRec.MedWages = Currency2String(W2FormRec(1).MedWages, 11) '12/03
    W2FormRec(1).STAWAGE = OldRound#(W2FormRec(1).STAWAGE + W2FormRec(1).BENFBOX1)
  End If
'*******************************

  For CntZZ = 0 To 51
    If Len(QPTrim$(W2SetUpRec(1).Deds(CntZZ).CHKDED)) And Deds(CntZZ) > 0 Then
      Select Case Left$(W2SetUpRec(1).Deds(CntZZ).CHKDED, 1)
      Case "P" 'pension
        W2FormRec(1).BOX15c = "X"
        W2RWRec.RetPlan = "1" '12/03
      Case "D" 'deferred compensation
        W2FormRec(1).BOX15c = "X"
        W2RWRec.RetPlan = "1" '12/03
      End Select
    End If
  Next
  If W2RWRec.RetPlan <> "1" Then W2RWRec.RetPlan = "0"
'  If QPTrim$(Emp2Rec(1).EmpLName) = "CAVINESS" Then Stop
  
  W2FormRec(1).BOX13AMT = 0
  W2FormRec(1).BOX13AM1 = 0
  W2FormRec(1).BOX13AM2 = 0
  W2FormRec(1).BOX13AM3 = 0
  W2FormRec(1).BOX14AMT = 0
  W2FormRec(1).BOX14AM1 = 0
  ThisLifeIns = 0
  ThisDefr401k = 0
  ThisDefr403b = 0
  ThisDefr408k6 = 0
  ThisDefr457b = 0
  ThisDefr501c18D = 0
  ThisNonStaStcks = 0
  ThisRoth401K = 0 'added for 2006
  
  For CntZZ = 0 To 51
    Select Case W2SetUpRec(1).Deds(CntZZ).AMTBOX
    Case "12a" '"13a"
      W2FormRec(1).BOX13AMT = OldRound#(W2FormRec(1).BOX13AMT + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TXt = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TXt) <> "" Then GoSub BOX13TXt
      End If
    Case "12b" '"13b"
      W2FormRec(1).BOX13AM1 = OldRound#(W2FormRec(1).BOX13AM1 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX1 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX1) <> "" Then GoSub BOX13TX1
      End If
    Case "12c" 'added Fall 04
      W2FormRec(1).BOX13AM2 = OldRound#(W2FormRec(1).BOX13AM2 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX2 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX2) <> "" Then GoSub BOX13TX2
      End If
    Case "12d" 'added Fall 04
      W2FormRec(1).BOX13AM3 = OldRound#(W2FormRec(1).BOX13AM3 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX13TX3 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX13TX3) <> "" Then GoSub BOX13TX3
      End If
    Case "14a"
      W2FormRec(1).BOX14AMT = OldRound#(W2FormRec(1).BOX14AMT + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX14TXT = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX14TXT) <> "" Then GoSub BOX14TXT
      End If
    Case "14b"
      W2FormRec(1).BOX14AM1 = OldRound#(W2FormRec(1).BOX14AM1 + Deds(CntZZ))
      If Deds(CntZZ) > 0 Then
        W2FormRec(1).BOX14TX1 = W2SetUpRec(1).Deds(CntZZ).DedCode
        If QPTrim$(W2FormRec(1).BOX14TX1) <> "" Then GoSub BOX14TX1
      End If
    End Select
  Next
  GoSub FinishIt
  Return
  
'  Exit Sub
  '12/03 from here to....
BOX13TXt: 'the W3 deferred compensation field only collects from Codes D-H and S
  '12/14/04 Discovered that the electronic fields were not being
  'collected properly if the same box was checked (say box 12a) but
  'different letters were assigned...changed to accumulate each
  'amount separately from the individual Deds(CntZZ) amounts instead of
  'from the W2FormRec(1) amounts
  Select Case QPTrim$(W2FormRec(1).BOX13TXt)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ) '12/14/04
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11) '12/14/04
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT 'added fall 04
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ)
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT '12/14/04
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ)
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT '12/14/04
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
    Case "S" 'added 8/24/04
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
    Case "AA" 'added for 2006
      ThisRoth401K = ThisRoth401K + Deds(CntZZ)
      W2RWRec.Roth401K = Currency2String(ThisRoth401K, 11)
      W2FormRec(1).W3DfCmp1 = W2FormRec(1).BOX13AMT
  End Select
  
Return

BOX13TX1:
  Select Case QPTrim$(W2FormRec(1).BOX13TX1)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ) '12/14/04
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11) '12/14/04
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ) '12/14/04
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
    Case "S"
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
    Case "AA" 'added for 2006
      ThisRoth401K = ThisRoth401K + Deds(CntZZ)
      W2RWRec.Roth401K = Currency2String(ThisRoth401K, 11)
      W2FormRec(1).W3DfCmp2 = W2FormRec(1).BOX13AM1
  End Select
  Return

BOX13TX2:
  Select Case QPTrim$(W2FormRec(1).BOX13TX2)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ)
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11)
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ) '12/14/04
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
    Case "S"
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
    Case "AA" 'added for 2006
      ThisRoth401K = ThisRoth401K + Deds(CntZZ)
      W2RWRec.Roth401K = Currency2String(ThisRoth401K, 11)
      W2FormRec(1).W3DfCmp3 = W2FormRec(1).BOX13AM2
  End Select
  Return

BOX13TX3:
  Select Case QPTrim$(W2FormRec(1).BOX13TX3)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ) '12/14/04
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11) '12/14/04
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ) '12/14/04
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
    Case "S"
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
    Case "AA" 'added for 2006
      ThisRoth401K = ThisRoth401K + Deds(CntZZ)
      W2RWRec.Roth401K = Currency2String(ThisRoth401K, 11)
      W2FormRec(1).W3DfCmp4 = W2FormRec(1).BOX13AM3
  End Select
  Return

BOX14TXT: 'according to 2006 W2 tax instructions (pg 5 for code AA) Roth 401K
'is to be reported only in box 12
  Select Case QPTrim$(W2FormRec(1).BOX14TXT)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ) '12/14/04
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11) '12/14/04
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ) '12/14/04
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
  End Select
  Return

BOX14TX1:
  Select Case QPTrim$(W2FormRec(1).BOX14TX1)
    Case "C"
      ThisLifeIns = ThisLifeIns + Deds(CntZZ) '12/14/04
      W2RWRec.LifeIns = Currency2String(ThisLifeIns, 11) '12/14/04
    Case "D"
      ThisDefr401k = ThisDefr401k + Deds(CntZZ) '12/14/04
      W2RWRec.Defr401k = Currency2String(ThisDefr401k, 11) '12/14/04
    Case "E"
      ThisDefr403b = ThisDefr403b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr403b = Currency2String(ThisDefr403b, 11) '12/14/04
    Case "F"
      ThisDefr408k6 = ThisDefr408k6 + Deds(CntZZ) '12/14/04
      W2RWRec.Defr408k6 = Currency2String(ThisDefr408k6, 11) '12/14/04
    Case "G"
      ThisDefr457b = ThisDefr457b + Deds(CntZZ) '12/14/04
      W2RWRec.Defr457b = Currency2String(ThisDefr457b, 11) '12/14/04
    Case "H"
      ThisDefr501c18D = ThisDefr501c18D + Deds(CntZZ) '12/14/04
      W2RWRec.Defr501c18D = Currency2String(ThisDefr501c18D, 11) '12/14/04
    Case "V"
      ThisNonStaStcks = ThisNonStaStcks + Deds(CntZZ) '12/14/04
      W2RWRec.NonStaStcks = Currency2String(ThisNonStaStcks, 11) '12/14/04
  End Select
  Return
  
FinishIt:
  If Len(QPTrim$(W2RWRec.LifeIns)) = 0 Then
    W2RWRec.LifeIns = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr401k)) = 0 Then
    W2RWRec.Defr401k = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr403b)) = 0 Then
    W2RWRec.Defr403b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr408k6)) = 0 Then
    W2RWRec.Defr408k6 = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr457b)) = 0 Then
    W2RWRec.Defr457b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr501c18D)) = 0 Then
    W2RWRec.Defr501c18D = ""
  End If
  If Len(QPTrim$(W2RWRec.NonStaStcks)) = 0 Then
    W2RWRec.NonStaStcks = ""
  End If
  If Len(QPTrim$(W2RWRec.Roth401K)) = 0 Then 'added for 2006
    W2RWRec.Roth401K = ""
  End If
  
  '...here
  Return

CheckEmp:
  If TPntr(0) Then            'if this emp has any transactions
    GoSub SumW2Info           'sum emp w2 info
    GoSub ApplyForm
  Else
    W2FormRec(1) = BW2FormRec(1)
    W2RWRec = BW2RWRec
  End If
Return

ExitW2SetUp:


End Sub

