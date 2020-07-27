Attribute VB_Name = "PR_Common"
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
  Public alnRpt$
  Public OutFileNames(1 To 20) As String
  Public InFileNames(1 To 20) As String
  Public ComputerName As String
  Public BadMaskFlag As Boolean
  Public EmpInfo(1 To 30) As String
  Public ToPrint1(1 To 10) As Integer
  Public ToPrint2(1 To 10) As Integer
  Public CurrCitiPath As String
  Public NewListFlag As Boolean
  Public StartPath As String
  Public NumOfAligns As Integer
  Public GlblQtr$ 'used in ESC report
  Public FundCnt4Rpt As Integer 'used in YTD Wage Distribution report
  Public DeductionSelNum As Integer 'used in Deduction report
  Public ThisRpt$ 'used in reprint report
  Public RptOpt As Integer 'used to determine the type of reports; graphic or text
  Public AccrualDate As Integer '12/12/02
  Public AccrualDateString$ '12/12/02
  Public ErrAcct() As String
  Public ErrAmt() As Double
  Public ErrType() As String
  Public GlobalCheckNum$ 'used solely for voiding a check
  Public GlobalTransNum As Double 'used solely for voiding a check
  Public GlobalName As String 'used solely for voiding a check
  Public ErrEmpNum$
  Public ErrCnt As Integer
  Public PayType As String
  Public OTRate As Double
  Public RegRate As Double
  Public ThisFreq As String
'  Public GEmpNum As Integer
  Public Twiddle As String
  Public FromPR As Boolean
  Public OverCnt As Integer
  Public bigName() As String
  
  Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
  Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
  lpBuffer As String, nSize As Long)
  
            Public Const PRData = "prdata\"
            Public Const GLData = "gldata\"
  
  Public Const StateTaxFileName = "PRSTATAX.DAT"
Public Const FederalTaxFileName = "PRFEDTAX.DAT"
   Public Const ErnCodeFileName = "PRERNCOD.DAT"
     Public Const LeaveFileName = "PRLEAVE.DAT"
       Public Const EICFileName = "PREICTBL.DAT"
      Public Const UnitFileName = "PRUNIT.DAT"
       Public Const SysFileName = "PRSYS.DAT"
 Public Const TransWorkFileName = "PRTRANST.DAT"
 Public Const TransHistFileName = "PRTRANSH.DAT"
    Public Const CheckPrintFile = "CHKPRNT.DAT"
      Public Const EmpData1Name = "PREMP1.DAT"
      Public Const EmpData2Name = "PREMP2.DAT"
      Public Const EmpData3Name = "PREMP3.DAT"
       Public Const EmpIdxLName = "PREMPL.IDX"  'name idx
       Public Const EmpIdxNName = "PREMPN.IDX"  'numb idx
    Public Const EMPNumFileName = "PREMPUNO.DAT"
    Public Const EMPPinFileName = "PREMPPIN.DAT"
    Public Const AccrueFileName = "PRACCRUE.DAT"
    Public Const ChecksFileName = "PRCHECKS.DAT"
   Public Const VoucherFileName = "PRVOUCHR.DAT"
 Public Const PPDefaultFileName = "PRPPDEF.DAT"
    Public Const RetireFileName = "PRRETIRE.DAT"
   Public Const DedCodeFileName = "PRDEDCOD.DAT"
   Public Const PRDraftFileName = "PRDRAFTI.DAT"
   Public Const EmpDataFileMask = "PRRPTS\PREMPRPT.DPM"
  Public Const PrinterSetUpFile = "PRPRNSET.DAT"
    Public Const GLAcctIdxFile = "BAACCTDX.DAT"
    Public Const JGLAcctIdxFile = "GLACCT.IDX"
      Public Const PRActiveFile = "PRDATA\PRACTIVE.FLG"    '*
      Public Const AcctFileName = "GLACCT.DAT"
      Public Const TransFileName = "GLTRANS.DAT"
      Public Const PPDraftInfoFileName = "PPDFINFO.DAT"
  Public Const DraftInfoFileName = "PRDATA\PRDRAFTI.DAT"
     Public Const Draft1FileName = "PRDATA\PRDRAFT1.DAT"
     Public Const Draft5FileName = "PRDATA\PRDRAFT5.DAT"
     Public Const Draft6FileName = "PRDATA\PRDRAFT6.DAT"
     Public Const Draft8FileName = "PRDATA\PRDRAFT8.DAT"
     Public Const Draft9FileName = "PRDATA\PRDRAFT9.DAT"
      Public Const ESCDataFileName = "PRDATA\PRESCCFG.DAT"
      Public Const TempAccrualName = "PRDATA\PRTMPACCRUAL.DAT" '12/11/02
      Public Const OldHistFileName = "prdata\oldtranH.dat"
      Public Const TempVoidFileName = "TEMPVOID.DAT"
      Public Const VoidChkPostName = "PRDATA\VOIDCKPST.DAT"
      Public Const PayRateName = "PAYRATE.DAT"
      Public Const PayRateIdxName = "PRDATA\PAYRTIDX.DAT"
      Public Const PayRateNumIdxName = "PRDATA\PYRTNUMIDX.DAT"
      Public Const DedAlertName = "PRDATA\DALERT.DAT"
      Public Const EarnAlertName = "PRDATA\EALERT.DAT"
      Public Const K401DedName = "PRDATA\401KDED.DAT"
      Public Const MessageName = "EMPMESS.DAT"
      Public Const OrbitEmpDataName = "OrbEmpData.Dat"
      Public Const W2ESubRA = "W2ESUBRA.DAT"
      Public Const GLFundFile = "GLFund.DAT"
      Public Const GlAcctFile = "GLAcct.DAT"
      Public Const GlBudgetTransFile = "BgtTrans.DAT"
      
'Public Sub Terminate2Shell()
'   Dim UBFrmCnt As Integer
'   ' Loop through the forms collection and unload each form.
'   Close
'   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
'       Unload Forms(UBFrmCnt)
'   Next
'   DoEvents
'   End
'End Sub
'Public Sub Terminate()
''   Dim UBFrmCnt As Integer
''
''   If PWcnt = -3 Then GoTo SSPW 'Southern Software Password
''   ' Loop through the forms collection and unload each form.
''   ClearInUsePRReg PWcnt 'we want this intact so if another user
''   'gets in payroll the "inuse" warning will pop up
''   Close
''SSPW:
''   For UBFrmCnt = Forms.Count - 1 To 0 Step -1
''      Unload Forms(UBFrmCnt)
''   Next
''   DoEvents
''   End
'End Sub
'
'Public Sub OpenAcctFile(AcctFileNum, Optional NumAccts As Integer)
'  Dim GLAcct As GLAcctRecType
'  Dim GLAcctRecLen As Integer
'  GLAcctRecLen = Len(GLAcct)
'  AcctFileNum = FreeFile
'  Open GLData + GlAcctFile For Random Shared As AcctFileNum Len = GLAcctRecLen
'  End Sub
'Public Sub OpenFundFile(GlFundHandle As Integer)
'  Dim GLFund As GLFundRecType
'  Dim GLFundRecLen As Integer
'  GLFundRecLen = Len(GLFund)
'  GlFundHandle = FreeFile
'  Open GLData + GLFundFile For Random Shared As GlFundHandle Len = GLFundRecLen
'End Sub
'Public Sub OpenW2ESubRA(W2ESubRAHandle As Integer)
'  Dim W2ESubRARec As W2ElectronicSubRA
'  Dim W2ESubRARecLen As Integer
'  W2ESubRARecLen = Len(W2ESubRARec)
'  W2ESubRAHandle = FreeFile
'  Open PRData + W2ESubRA For Random Shared As W2ESubRAHandle Len = W2ESubRARecLen
'End Sub
'Public Sub OpenOrbEmpDataFile(OrbHandle As Integer)
'  Dim OrbLen As Integer
'  Dim OrbRec As OrbitEmpData
'  OrbLen = Len(OrbRec)
'  OrbHandle = FreeFile
'  Open PRData + OrbitEmpDataName For Random Shared As OrbHandle Len = OrbLen
'End Sub
'Public Sub OpenEmpMessage(MessHandle As Integer)
'  Dim MessLen As Integer
'  Dim MessRec As PRMessRecType
'  MessLen = Len(MessRec)
'  MessHandle = FreeFile
'  Open PRData + MessageName For Random Shared As MessHandle Len = MessLen
'End Sub
'Public Sub OpenGLTransFile(GlTransHandle As Integer)
'  Dim GLTransRec As GLTransRecType
'  Dim GLTransRecLen As Integer
'  GLTransRecLen = Len(GLTransRec)
'  GlTransHandle = FreeFile
'  Open GLData + TransFileName For Random Shared As GlTransHandle Len = GLTransRecLen
'End Sub
'Public Sub OpenGLBudgetTransFile(GlBudgetTransHandle As Integer)
'  Dim GLBudgetTransRec As GLTransRecType
'  Dim GLBudgetTransRecLen As Integer
'  GLBudgetTransRecLen = Len(GLBudgetTransRec)
'  GlBudgetTransHandle = FreeFile
'  Open GLData + GlBudgetTransFile For Random Shared As GlBudgetTransHandle Len = GLBudgetTransRecLen
'End Sub
'Public Sub Open401KDedFile(K401DedHandle As Integer)
'  Dim K401DedLen As Integer
'  Dim K401Ded As K401DedType
'  K401DedLen = Len(K401Ded)
'  K401DedHandle = FreeFile
'  Open K401DedName For Random Shared As K401DedHandle Len = K401DedLen
'End Sub
'Public Sub OpenEarnAlertFile(EarnAlertHandle As Integer)
'  Dim EarnAlertLen As Integer
'  Dim EarnAlert As TempEarnAlertType
'  EarnAlertLen = Len(EarnAlert)
'  EarnAlertHandle = FreeFile
'  Open EarnAlertName For Random Shared As EarnAlertHandle Len = EarnAlertLen
'End Sub
'Public Sub OpenDedAlertFile(DedAlertHandle As Integer)
'  Dim DedAlertLen As Integer
'  Dim DedAlert As TempDedAlertType
'  DedAlertLen = Len(DedAlert)
'  DedAlertHandle = FreeFile
'  Open DedAlertName For Random Shared As DedAlertHandle Len = DedAlertLen
'End Sub
'Public Sub OpenPayRateIdxFile(PayRateIdxHandle As Integer)
'  Dim PayRateIdxLen As Integer
'  Dim PayRateIdx As PayRateIndexType
'  PayRateIdxLen = Len(PayRateIdx)
'  PayRateIdxHandle = FreeFile
'  Open PayRateIdxName For Random Shared As PayRateIdxHandle Len = PayRateIdxLen
'End Sub
'
'Public Sub OpenPayRateNumIdxFile(PayRateIdxHandle As Integer)
'  Dim PayRateIdxLen As Integer
'  Dim PayRateIdx As PayRateIdxNumType
'  PayRateIdxLen = Len(PayRateIdx)
'  PayRateIdxHandle = FreeFile
'  Open PayRateNumIdxName For Random Shared As PayRateIdxHandle Len = PayRateIdxLen
'End Sub
'
'Public Sub OpenPayRateFile(PayRateHandle As Integer)
'  Dim PayRateLen As Integer
'  Dim PayRate As PayRateType
'  PayRateLen = Len(PayRate)
'  PayRateHandle = FreeFile
'  Open PRData + PayRateName For Random Shared As PayRateHandle Len = PayRateLen
'End Sub
'Public Sub OpenTempVoidFile(TempVoidHandle As Integer) '12/12/02
'  Dim TempVoidLen As Integer
'  Dim TempVoid As VoidCheckType
'  TempVoidLen = Len(TempVoid)
'  TempVoidHandle = FreeFile
'  Open PRData + TempVoidFileName For Random Shared As TempVoidHandle Len = TempVoidLen
'End Sub
'Public Sub OpenVoidChkPostFile(VoidPostHandle As Integer) '12/12/02
'  Dim VoidPostLen As Integer
'  Dim VoidPost As VoidCheckType
'  VoidPostLen = Len(VoidPost)
'  VoidPostHandle = FreeFile
'  Open VoidChkPostName For Random Shared As VoidPostHandle Len = VoidPostLen
'End Sub
'Public Sub OpenAccrualDatesFile(AccrualDatesHandle As Integer) '12/12/02
'  Dim AccrualDatesLen As Integer
'  Dim AccrualDates As TempAccrualType
'  AccrualDatesLen = Len(AccrualDates)
'  AccrualDatesHandle = FreeFile
'  Open PRData + AccrueFileName For Random Shared As AccrualDatesHandle Len = AccrualDatesLen
'End Sub
'Public Sub OpenTempAccrualFile(TempAccrualHandle As Integer) '12/11/02
'  Dim TempAccrualLen As Integer
'  Dim TempAccrual As TempAccrualType
'  TempAccrualLen = Len(TempAccrual)
'  TempAccrualHandle = FreeFile
'  Open TempAccrualName For Random Shared As TempAccrualHandle Len = TempAccrualLen
'End Sub
'Public Sub OpenOSChekFile(OSChekFileNum, NumOSChks)
'  Dim OSChekLen As Integer
'  Dim OSChek    As OSChkRecType
'  OSChekLen = Len(OSChek)
'  OSChekFileNum = FreeFile
'  Open "crchek.dat" For Random Shared As OSChekFileNum Len = OSChekLen
'  NumOSChks = LOF(OSChekFileNum) \ (OSChekLen)
'End Sub
'
'Public Sub OpenPPDraftInfo(PPDraftInfoHandle As Integer)
'  Dim PPDraftInfoRec As GLAcctIndexType
'  Dim PPDraftInfoRecLen As Integer
'  PPDraftInfoRecLen = Len(PPDraftInfoRec)
'  PPDraftInfoHandle = FreeFile
'  Open PPDraftInfoFileName For Random Shared As PPDraftInfoHandle Len = PPDraftInfoRecLen
'End Sub
'
'Public Sub OpenGLAcctIdx(GLAcctIdxHandle As Integer)
'  Dim GLAcctRec As GLAcctIndexType
'  Dim GLAcctRecLen As Integer
'  GLAcctRecLen = Len(GLAcctRec)
'  GLAcctIdxHandle = FreeFile
'  Open GetCitiDirFolder + JGLAcctIdxFile For Random Shared As GLAcctIdxHandle Len = GLAcctRecLen
'End Sub
'Public Sub OpenGLAcctFile(GlHandle As Integer)
'  Dim GLRec As GLAcctRecType
'  Dim GLRecLen As Integer
'  GLRecLen = Len(GLRec)
'  GlHandle = FreeFile
'  Open GetCitiDirFolder + AcctFileName For Random Shared As GlHandle Len = GLRecLen
'End Sub
'
'Public Sub OpenChecksFile(ChecksHandle As Integer)
'  Dim ChecksRec As PRCheckRecType
'  Dim ChecksRecLen As Integer
'  ChecksRecLen = Len(ChecksRec)
'  ChecksHandle = FreeFile
'  Open PRData + ChecksFileName For Random Shared As ChecksHandle Len = ChecksRecLen
'End Sub
'
'Public Sub OpenEmpNumFile(EmpNumHandle As Integer)
'  Dim EmpNumRec As EmpNumType
'  Dim EmpNumRecLen As Integer
'  EmpNumRecLen = Len(EmpNumRec)
'  EmpNumHandle = FreeFile
'  Open PRData + EMPNumFileName For Random Shared As EmpNumHandle Len = EmpNumRecLen
'End Sub
'
'Public Sub OpenPRChecksFile(PRChecksHandle As Integer)
'  Dim PRChecksRec As PRCheckRecType
'  Dim PRChecksRecLen As Integer
'  PRChecksRecLen = Len(PRChecksRec)
'  PRChecksHandle = FreeFile
'  Open PRData + ChecksFileName For Random Shared As PRChecksHandle Len = PRChecksRecLen
'End Sub
'
'Public Sub OpenPPDefaultFile(PPDefaultHandle As Integer)
'  Dim PPDefaultRec As PeriodDefaultRecType
'  Dim PPDefaultRecLen As Integer
'  PPDefaultRecLen = Len(PPDefaultRec)
'  PPDefaultHandle = FreeFile
'  Open PRData + PPDefaultFileName For Random Shared As PPDefaultHandle Len = PPDefaultRecLen
'End Sub
'
''Public Sub OpenPrinterSetupFile(PrinterSUFHandle As Integer)
'''  Dim PrinterSUFRec As PRNSetupRecType
'''  Dim PrinterSUFRecLen As Integer
'''  PrinterSUFRecLen = Len(PrinterSUFRec)
'''  PrinterSUFHandle = FreeFile
'''  Open PRData + PrinterSetUpFile For Random Shared As PrinterSUFHandle Len = PrinterSUFRecLen
''End Sub
'
'Public Sub OpenLeaveFileName(LeaveHandle As Integer)
'  Dim LeaveRec As LeaveRecType
'  Dim LeaveRecLen As Integer
'  LeaveRecLen = Len(LeaveRec)
'  LeaveHandle = FreeFile
'  Open PRData + LeaveFileName For Random Shared As LeaveHandle Len = LeaveRecLen
'End Sub
'
'Public Sub OpenStateTaxFileName(STFNHandle As Integer)
'  Dim STFNRec As StateTaxRecType
'  Dim STFNRecLen As Integer
'  STFNRecLen = Len(STFNRec)
'  STFNHandle = FreeFile
'  Open PRData + StateTaxFileName For Random Shared As STFNHandle Len = STFNRecLen
'End Sub
'
'Public Sub OpenTransWorkFile(TransWorkFileHandle As Integer)
'  Dim TransWorkFileRec As TransRecType
'  Dim TransWorkRecLen As Integer
'  TransWorkRecLen = Len(TransWorkFileRec)
'  TransWorkFileHandle = FreeFile
'  Open PRData + TransWorkFileName For Random Shared As TransWorkFileHandle Len = TransWorkRecLen
'End Sub
'
'Public Sub OpenTransHistFile(TransHistFileHandle As Integer)
'  Dim TransHistFileRec As TransRecType
'  Dim TransHistRecLen As Integer
'  TransHistRecLen = Len(TransHistFileRec)
'  TransHistFileHandle = FreeFile
'  Open PRData + TransHistFileName For Random Shared As TransHistFileHandle Len = TransHistRecLen
'End Sub
''****************************************************************************
''OldRounds a double precision value to nearest hundredth
''****************************************************************************

Public Function OldRound#(N As Double)
  OldRound# = Int(N * 100 + 0.50000001) / 100
End Function
'
'Public Sub OpenEmpIdxLNameFile(EmpIdxLNameHandle As Integer)
'  EmpIdxLNameHandle = FreeFile
'  Open PRData + EmpIdxLName For Random Shared As EmpIdxLNameHandle Len = 2
'End Sub
'
'Public Sub OpenEmpIdxNNameFile(EmpIdxNNameHandle As Integer)
'  EmpIdxNNameHandle = FreeFile
'  Open PRData + EmpIdxNName For Random Shared As EmpIdxNNameHandle Len = 2
'End Sub
'Public Sub OpenEmpData3File(EmpData3FileHandle As Integer)
'  Dim EmpData3FileRec As EmpData3Type
'  Dim EmpData3RecLen As Integer
'  EmpData3RecLen = Len(EmpData3FileRec)
'  EmpData3FileHandle = FreeFile
'  Open PRData + EmpData3Name For Random Shared As EmpData3FileHandle Len = EmpData3RecLen
'End Sub
'Public Sub OpenEmpData2File(EmpData2FileHandle As Integer)
'  Dim EmpData2FileRec As EmpData2Type
'  Dim EmpData2RecLen As Integer
'  EmpData2RecLen = Len(EmpData2FileRec)
'  EmpData2FileHandle = FreeFile
'  Open PRData + EmpData2Name For Random Shared As EmpData2FileHandle Len = EmpData2RecLen
'End Sub
'
'Public Sub OpenEmpData1File(EmpData1FileHandle As Integer)
'  Dim EmpData1FileRec As EmpData1Type
'  Dim EmpData1RecLen As Integer
'  EmpData1RecLen = Len(EmpData1FileRec)
'  EmpData1FileHandle = FreeFile
'  Open PRData + EmpData1Name For Random Shared As EmpData1FileHandle Len = EmpData1RecLen
'End Sub
'Public Sub OpenStateTaxFile(StateTaxFileHandle As Integer)
'  Dim StateTaxFileRec As StateTaxRecType
'  Dim StateTaxRecLen As Integer
'  StateTaxRecLen = Len(StateTaxFileRec)
'  StateTaxFileHandle = FreeFile
'  Open PRData + StateTaxFileName For Random Shared As StateTaxFileHandle Len = StateTaxRecLen
'End Sub
'
'Public Sub OpenFedTaxFile(FedTaxFileHandle As Integer)
'  Dim FedTaxFileRec As FederalTaxRecType
'  Dim FedTaxRecLen As Integer
'  FedTaxRecLen = Len(FedTaxFileRec)
'  FedTaxFileHandle = FreeFile
'  Open PRData + FederalTaxFileName For Random Shared As FedTaxFileHandle Len = FedTaxRecLen
'End Sub
'
'Public Sub OpenDedCodeFile(DedCodeFileHandle As Integer)
'  Dim DedCodeFileRec As DedCodeRecType
'  Dim DedCodeRecLen As Integer
'  DedCodeRecLen = Len(DedCodeFileRec)
'  DedCodeFileHandle = FreeFile
'  Open PRData + DedCodeFileName For Random Shared As DedCodeFileHandle Len = DedCodeRecLen
'End Sub
'
'Public Sub OpenErnCodeFile(ErnCodeFileHandle As Integer)
'  Dim ErnCodeFileRec As ErnCodeRecType
'  Dim ErnCodeRecLen As Integer
'  ErnCodeRecLen = Len(ErnCodeFileRec)
'  ErnCodeFileHandle = FreeFile
'  Open PRData + ErnCodeFileName For Random Shared As ErnCodeFileHandle Len = ErnCodeRecLen
'End Sub
'
'Public Sub OpenPRDraftFile(PRDraftFileHandle As Integer)
'  Dim PRDraftFileRec As DraftInfoFileName
'  Dim PRDraftRecLen As Integer
'  PRDraftRecLen = Len(PRDraftFileRec)
'  PRDraftFileHandle = FreeFile
'  Open PRData + PRDraftFileName For Random Shared As PRDraftFileHandle Len = PRDraftRecLen
'End Sub
'
'Public Sub OpenRetFile(RetFileHandle As Integer)
'  Dim RetFileRec As RetireRecType
'  Dim RetRecLen As Integer
'  RetRecLen = Len(RetFileRec)
'  RetFileHandle = FreeFile
'  Open PRData + RetireFileName For Random Shared As RetFileHandle Len = RetRecLen
'End Sub
'
'Public Sub OpenEICFile(EICFileHandle As Integer)
'  Dim EICFileRec As EICRecType
'  Dim EICRecLen As Integer
'  EICRecLen = Len(EICFileRec)
'  EICFileHandle = FreeFile
'  Open PRData + EICFileName For Random Shared As EICFileHandle Len = EICRecLen
'End Sub
'
'Public Sub OpenSysFile(SysFileHandle As Integer)
'  Dim SysFileRec As RegDSysFileRecType
'  Dim SysRecLen As Integer
'  SysRecLen = Len(SysFileRec)
'  SysFileHandle = FreeFile
'  Open PRData + SysFileName For Random Shared As SysFileHandle Len = SysRecLen
'End Sub
'
'Public Sub OpenUnitFile(FileHandle As Integer)
'  Dim UnitFileRec As UnitFileRecType
'  Dim UnitRecLen As Integer
'  UnitRecLen = Len(UnitFileRec)
'  FileHandle = FreeFile
'  Open PRData + UnitFileName For Random Shared As FileHandle Len = UnitRecLen
'End Sub
'
Public Function QPTrim$(Text As String)
  Dim StrLen As Long
  Dim Cnt As Long
  Dim thischar As Integer
  StrLen = Len(Text)
  For Cnt = 1 To StrLen
    thischar = Asc(Mid$(Text, Cnt, 1))
    If thischar = 0 Then
      Mid$(Text$, Cnt, 1) = " "
    End If
  Next
  QPTrim$ = Trim$(Text)
End Function

'Public Function ReplaceString$(Text As String, ChangeThis As String, ToThis As String)
'  Dim StrLen As Long
'  Dim Cnt As Long
'  Dim NewText As String
'  Dim thischar$
'  Dim CTChar$
'  Dim TTChar$
'  Dim CTLen As Integer
'  Dim TTLen As Integer
'  Dim BigLen As Integer
'  'this function takes the incoming text and rebuilds it one
'  'letter at a time until it encounters the text to change
'  'at which time it replaces the text to change with the
'  'new text
'  StrLen = Len(Text)
'  CTLen = Len(ChangeThis$)
'  TTLen = Len(ToThis$)
'  If CTLen > TTLen Then
'    BigLen = CTLen
'  ElseIf TTLen > CTLen Then
'    BigLen = TTLen
'  Else
'    BigLen = CTLen
'  End If
'
'  For Cnt = 1 To StrLen 'set up loop to iterate thru entire text
'    thischar = Mid$(Text, Cnt, 1) 'step thru text a letter at a time
'    CTChar = Mid$(Text, Cnt, CTLen) 'starting with the current letter
'    'read ahead the length of the text "change this"
'    If CTChar = ChangeThis Then 'if we find the "change this" in the
'    'text
'      NewText = NewText + ToThis 'assign the length of CTChar to "ToThis"
'      'inside the rebuilt new text
'      Cnt = Cnt + BigLen - 1 'advance count to compensate for the addition of
'      'CTChar
'    Else
'      NewText = NewText + thischar 'build new text one letter at a time
'    End If
'  Next
'  ReplaceString$ = Trim$(NewText) 'rim out the new text
'  Text = ReplaceString$ 'old text is now new text
'End Function

Public Sub KillFile(FileName As String)
  On Local Error Resume Next
  If Exist(FileName$) Then 'added 7/24
    Kill FileName$
  End If
  On Error GoTo 0
End Sub

''Public Function digitCheck$(Text As String) 'commented out 8/15
''  Dim StrLen As Long
''  Dim cnt As Long
''  Dim ThisChar As Integer
''  'this function traps for digits like 3.456.78 where the user
''  'inserts an unwanted decimal which causes errors
''  digitCheck = True
''  StrLen = Len(Text)
''  For cnt = 1 To StrLen
''    ThisChar = Asc(Mid$(Text, cnt, 1))
''    If ThisChar < 48 Or ThisChar > 57 Then
''      digitCheck = False
''      Exit For
''    End If
''  Next
''End Function
''Public Function PromptVoid(frm As Form) As SaveChangeOptions1
''  'all "prompt" functions work in conjunction with a warning
''  'screen that requires an option to be selected by the user
''  frmWarnVoidThisChk.Show vbModal, frm
''  PromptVoid = frmWarnVoidThisChk.Selection
''  Unload frmWarnVoidThisChk
''End Function
''Public Function PromptPRTRemove(frm As Form) As SaveChangeOptions1
''  frmWarnRemovePRT.Show vbModal, frm
''  PromptPRTRemove = frmWarnRemovePRT.Selection
''  Unload frmWarnRemovePRT
''End Function
''Public Function PromptPRTAccrue(frm As Form) As SaveChangeOptions1 '8/7
''  frmWarnAccrueNow.Show vbModal, frm '8/7
''  PromptPRTAccrue = frmWarnAccrueNow.Selection '8/7
''  Unload frmWarnAccrueNow '8/7
''End Function
''Public Function PromptUserAlreadyActive(frm As Form) As PRInUse
''  frmWarnInUse.Show vbModal, frm
''  PromptUserAlreadyActive = frmWarnInUse.Selection
'''  Unload frmWarnPRInProgress
''End Function
''
''Public Function PromptPRInProgress(frm As Form) As SaveChangeOptions1
''  frmWarnPRInProgress.Show vbModal, frm
''  PromptPRInProgress = frmWarnPRInProgress.Selection
'''  Unload frmWarnPRInProgress
''End Function
''
''Public Function PromptSaveChanges(frm As Form) As SaveChangeOptions1
''  DoEvents
''  frmChangedWarning.Show vbModal, frm
''  PromptSaveChanges = frmChangedWarning.Selection
''  Unload frmChangedWarning
''End Function
''
''Public Function PromptBadSSNNum(frm As Form) As SaveChangeOptions1
''  frmWarnBadSSN.Show vbModal, frm
''  PromptBadSSNNum = frmWarnBadSSN.Selection
''  Unload frmWarnBadSSN
''End Function
''Public Function PromptBadGLNum(frm As Form) As SaveChangeOptions1
''  frmWarnBadGLNum.Show vbModal, frm
''  PromptBadGLNum = frmWarnBadGLNum.Selection
''  Unload frmWarnBadGLNum
''End Function
''Public Function PromptBadGLNumVer2(frm As Form) As SaveChangeOptions1
''  frmWarnBadGLNumVer2.Show vbModal, frm
''  PromptBadGLNumVer2 = frmWarnBadGLNumVer2.Selection
''  Unload frmWarnBadGLNumVer2
''  DoEvents
''End Function
''
''Public Function PromptSaveLvBnftChanges(frm As Form) As SaveChangeOptions1
''  frmWarningLvBnft.Show vbModal, frm
''  PromptSaveLvBnftChanges = frmWarningLvBnft.Selection
''  Unload frmWarningLvBnft
''End Function
'
'Public Function CheckValDate(ValCheck As String)
'  Dim Month As Integer, Day As Integer, Year As Integer
'  Month = Val(Mid(ValCheck, 1, 2))
'  Day = Val(Mid(ValCheck, 4, 2))
'  Year = Val(Mid(ValCheck, 7, 4))
'  'Checks date if Blank then won't check for valid date
'  'and then checks each section, month, day and year
'  'if any section wrong then returns false value
'      If InStr(ValCheck, "_") <= 0 Then
'          If ((Month > 0) And (Month < 13)) Then
'              If Day > 0 And Day < 32 Then
'                  If Year > 1919 And Year < 2099 Then
'                      CheckValDate = True
'                  End If
'              End If
'          End If
'      End If
'End Function
'
''This function is a replacement for the QuickPak FileSize function.
''Due to the way Windows NT updates a file's size in the directory, an
''error can occur using DOS Function 4Eh (Find first file service) to
''read a file's size from the Directory. You can force Windows NT to
''commit the directory info by just opening the file again.
'Public Function FileSize(FileName$) As Long
'  Dim FileHandle As Integer
'  If Exist(FileName$) Then
'    FileHandle = FreeFile
'    Open FileName$ For Binary As FileHandle
'    FileSize = LOF(FileHandle)
'    Close FileHandle
'  Else
'    FileSize = 0
'  End If
'End Function
'
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
  On Error GoTo 0
End Function
'
''Public Sub ViewPrint(ReportFile As String, Title As String, Optional ForceSBar As Boolean, Optional PgNum As Integer, Optional Algn As Boolean, Optional AlgnRptfile As String)
''   frmLoadingRpt.Show
''   frmViewPrint.ReportName = ReportFile$
''   frmViewPrint.Caption = Title
''   frmViewPrint.PgNum = PgNum
''   frmViewPrint.cmdAlignment.Visible = False
''   If ForceSBar Then
''     frmViewPrint.fpMemo1.ScrollBars = BothFixed
''   Else
''     frmViewPrint.fpMemo1.ScrollBars = BothAuto
''   End If
''   If Algn Then
''     frmViewPrint.cmdAlignment.Enabled = True
''     frmViewPrint.AlignRpt = AlgnRptfile$
''    Else
''      frmViewPrint.cmdAlignment.Enabled = False
''    End If
''   frmViewPrint.Show 1
''   Unload frmLoadingRpt
''   doAlign = False
''End Sub

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


'Public Sub MakeEmpIndexs()
'  Dim EmpData2FileHandle As Integer, NumOfEmpRec As Integer
'  Dim EmpData2FileRec As EmpData2Type
'  Dim x As Integer, IdxHandle As Integer
'  ReDim EmpIdxNNameRec(1 To 1) As NumbSortIdxType
'  ReDim EmpIdxLNameRec(1 To 1) As NameSortIdxType
'  Dim TempIdxRec As NumbSortIdxType
'  Dim TempLIdxRec As NameSortIdxType
'  Dim OutOfOrder  As Boolean, Found As Integer
'  'this function takes the current list of employees and
'  'sorts them by number and last name to an index for each
'  OpenEmpData2File EmpData2FileHandle
'  NumOfEmpRec = LOF(EmpData2FileHandle) / Len(EmpData2FileRec)
'  For x = 1 To NumOfEmpRec
'    Get EmpData2FileHandle, x, EmpData2FileRec
'    Found = Found + 1
'    ReDim Preserve EmpIdxNNameRec(1 To Found) As NumbSortIdxType
'    'ReDim to increase the array size as new employees are added
'    ReDim Preserve EmpIdxLNameRec(1 To Found) As NameSortIdxType
'    EmpIdxLNameRec(Found).EmpName = QPTrim$(EmpData2FileRec.EmpLName)
'    EmpIdxLNameRec(Found).DataRecNum = x 'add data to last name index
'    RSet EmpIdxNNameRec(Found).EmpNumb = QPTrim$(EmpData2FileRec.EmpNo)
'    EmpIdxNNameRec(Found).DataRecNum = x 'add data to number index
'  Next x
'  Close EmpData2FileHandle
'
'  Do  'Sort the employee numbers
'    OutOfOrder = False          'assume it's sorted
'    For x = 1 To Found - 1
'      If Val(EmpIdxNNameRec(x).EmpNumb) > Val(EmpIdxNNameRec(x + 1).EmpNumb) Then
'        LSet TempIdxRec = EmpIdxNNameRec(x)
'        LSet EmpIdxNNameRec(x) = EmpIdxNNameRec(x + 1)
'        LSet EmpIdxNNameRec(x + 1) = TempIdxRec
'        OutOfOrder = True       'we're not done yet
'      End If
'    Next
'  Loop While OutOfOrder
'
'  Do  'Sort the employee Names
'    OutOfOrder = False          'assume it's sorted
'    For x = 1 To Found - 1
'      If EmpIdxLNameRec(x).EmpName > EmpIdxLNameRec(x + 1).EmpName Then
'        LSet TempLIdxRec = EmpIdxLNameRec(x)
'        LSet EmpIdxLNameRec(x) = EmpIdxLNameRec(x + 1)
'        LSet EmpIdxLNameRec(x + 1) = TempLIdxRec
'        OutOfOrder = True       'we're not done yet
'      End If
'    Next
'  Loop While OutOfOrder
'
'  'delete any existing number index
'  If Exist(PRData + EmpIdxNName) Then
'    Kill PRData + EmpIdxNName
'  End If
'
'  'delete any existing last name index
'  If Exist(PRData + EmpIdxLName) Then
'    Kill PRData + EmpIdxLName
'  End If
'
'  'build new number index
'  OpenEmpIdxNNameFile IdxHandle
'  For x = 1 To Found
'    Put IdxHandle, x, EmpIdxNNameRec(x).DataRecNum
'  Next
'  Close IdxHandle
'
'  'build new last name index
'  OpenEmpIdxLNameFile IdxHandle
'  For x = 1 To Found
'    Put IdxHandle, x, EmpIdxLNameRec(x).DataRecNum
'  Next
'  Close IdxHandle
'End Sub
'Public Function Date2Num%(TheDate$)
' 'useful function throughout program...
' 'takes a string date and converts into a number based on 12/31/1979
'  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
'End Function
'Public Function MakeMDYDate(ByRef DateToFix As String) As String
'  Dim CCYY As String
'  Dim MM As String
'  Dim DD As String
'
'  CCYY = Mid(DateToFix, 1, 4)
'  MM = Mid(DateToFix, 5, 2)
'  DD = Mid(DateToFix, 7, 2)
'  MakeMDYDate$ = MM + "/" + DD + "/" + CCYY
'End Function
'
'
Public Function MakeRegDate(ByVal DateNumb)
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function
'Public Function MakeRegDateDash(ByVal DateNumb)
'  Dim Month As Integer, ThisDate As String
'  'function does the opposite of Date2Num
'  If DateNumb = -32767 Then
'    MakeRegDateDash = "%%%%%%%%%% "
'  Else
'    MakeRegDateDash = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm-dd-yyyy")
'  End If
'End Function
'Public Function CheckCitiDir(CityPakDir$) As Integer
'  On Local Error GoTo Oh_Shucks
'  Dim Handle As Integer
'  'function verifies a valid path to the Citipak directory
'  CityPakDir$ = QPTrim$(CityPakDir$)
'
'  If Len(QPTrim$(CityPakDir$)) = 0 Then 'path not saved yet
'    CheckCitiDir = 1
'    Exit Function
'  End If
'
'  If Right$(CityPakDir$, 1) <> "\" Then 'adds a back slash if absent
'    CityPakDir$ = CityPakDir$ + "\"
'  End If
'
'  Handle = FreeFile
'  'If the next open statement cannot occur without an error then
'  'the On Local Error statement above sends the function back as false
'  Open CityPakDir$ + "TestCitiDir.TXT" For Binary As Handle
'  'if we get here then all is well
'  Close
'  CheckCitiDir = True
'  Exit Function
'
'Oh_Shucks:
'  CheckCitiDir = False
'  Close
'
'End Function
'Public Function AddDashToSSN(ByVal SSN)
'  Dim NewSSN As String
'  'function serves as a method for converting a social security
'  'number without dashed back to one with dashed before loading
'  'on a screen
'  If Mid(SSN, 4, 1) <> "-" And Mid(SSN, 7, 1) <> "-" Then
'      NewSSN = Mid(SSN, 1, 3) + "-" + Mid(SSN, 4, 2) + "-" + Mid(SSN, 6, 4)
'      AddDashToSSN = NewSSN
'  Else
'      AddDashToSSN = SSN
'  End If
'End Function
'
'Public Function MonthName$(ByVal MonthNo As Integer)
'  'Used by a couple of reports (Retirement and SEPP)
'  Select Case MonthNo
'  Case 1
'    MonthName$ = "January"
'  Case 2
'    MonthName$ = "February"
'  Case 3
'    MonthName$ = "March"
'  Case 4
'    MonthName$ = "April"
'  Case 5
'    MonthName$ = "May"
'  Case 6
'    MonthName$ = "June"
'  Case 7
'    MonthName$ = "July"
'  Case 8
'    MonthName$ = "August"
'  Case 9
'    MonthName$ = "September"
'  Case 10
'    MonthName$ = "October"
'  Case 11
'    MonthName$ = "November"
'  Case 12
'    MonthName$ = "December"
'  End Select
'
'End Function
'
'Public Sub CreateEmpTransRecs(RecNo As Long)
'  Dim Emp2Handle As Integer
'  Dim Emp2Rec(1) As EmpData2Type
'  Dim TempTransRec(1) As TransRecType
'  Dim PayType$, PayFreq$
'  Dim PHandle As Integer
'  Dim PPDRec As PeriodDefaultRecType
'  Dim x As Integer
'  Dim TotSalDist As Double
'  Dim Fund As Integer
'  Dim Dept As Integer
'  Dim Detail As Integer
'
'  '10/19/04 discovered that with non-split calculations the program
'  'was producing separate wage results for the same GL number
'  'if some numbers were entered with the dasg and others were
'  'entered without dashes. All 10/19/04 code was inserted to make
'  'all numbers the same by adding dashes where they were needed.
'
'  Call GetAcctStruct(CurrCitiPath$, Fund, Dept, Detail) '10/19/04
'  'this sub screens out employees that won't be included in this pay
'  'period then loads each employee's current transaction data to
'  'TempTransRec(1) and saved after CalcPay concludes
'
'  OpenEmpData2File Emp2Handle
'  Get Emp2Handle, RecNo, Emp2Rec(1)
'  Close Emp2Handle
'  'screen out terminated or deleted employees
'  If Emp2Rec(1).EMPTDATE <> 0 Or Emp2Rec(1).Deleted = True Then
'    Exit Sub
'  End If
'
'  OpenPPDefaultFile PHandle
'  Get PHandle, 1, PPDRec
'  Close PHandle
'
'  TempTransRec(1).TActive = True 'denotes this employee is included
'  'in this pay period
'  PayType$ = UCase$(QPTrim$(Emp2Rec(1).EMPPTYPE))
'  PayFreq$ = UCase$(QPTrim$(Emp2Rec(1).EMPPFREQ))
'  TempTransRec(1).EmpPin = Emp2Rec(1).EmpPin
'  TempTransRec(1).PayPdStart = PPDRec.PERBEG
'  TempTransRec(1).PayPdEnd = PPDRec.PEREND
'
'  If Emp2Rec(1).EMPPRATE < 0 Then 'screens out errors in pay rates
'    TempTransRec(1).BaseRate = 0
'  Else
'    TempTransRec(1).BaseRate = Emp2Rec(1).EMPPRATE
'  End If
'
'  If Emp2Rec(1).EMPORATE < 0 Then 'screens out errors in overtime rates
'    TempTransRec(1).OTRate = 0
'  Else
'    TempTransRec(1).OTRate = Emp2Rec(1).EMPORATE
'  End If
'
'  'if statement looks for add'l earnings that will be included in this
'  'pay period as set in the PR Default screen and then compared
'  'with whether or not this employee has a value saved for each
'  If PPDRec.USEAE1 = "Y" Then
'    If Emp2Rec(1).EMPEAMT1 < 0 Then
'      TempTransRec(1).EAmt(1) = 0
'    Else
'      TempTransRec(1).EAmt(1) = Emp2Rec(1).EMPEAMT1
'    End If
'    TempTransRec(1).EDist(1).EAmt = TempTransRec(1).EAmt(1)
'    If QPTrim$(Emp2Rec(1).EMPEACT1) <> "" Then '10/19/04
'      Emp2Rec(1).EMPEACT1 = AddDashesToGLNumber(Emp2Rec(1).EMPEACT1, Fund, Dept, Detail)
'    End If
'    TempTransRec(1).EDist(1).EAcct = Emp2Rec(1).EMPEACT1
'  End If
'  If PPDRec.USEAE2 = "Y" Then
'    If Emp2Rec(1).EMPEAMT2 < 0 Then
'      TempTransRec(1).EAmt(2) = 0
'    Else
'      TempTransRec(1).EAmt(2) = Emp2Rec(1).EMPEAMT2
'    End If
'    TempTransRec(1).EDist(2).EAmt = TempTransRec(1).EAmt(2)
'    If QPTrim$(Emp2Rec(1).EMPEACT2) <> "" Then '10/19/04
'      Emp2Rec(1).EMPEACT2 = AddDashesToGLNumber(Emp2Rec(1).EMPEACT2, Fund, Dept, Detail)
'    End If
'    TempTransRec(1).EDist(2).EAcct = Emp2Rec(1).EMPEACT2
'  End If
'
'  If PPDRec.USEAE3 = "Y" Then
'    If Emp2Rec(1).EMPEAMT3 < 0 Then
'      TempTransRec(1).EAmt(3) = 0
'    Else
'      TempTransRec(1).EAmt(3) = Emp2Rec(1).EMPEAMT3
'    End If
'    TempTransRec(1).EDist(3).EAmt = TempTransRec(1).EAmt(3)
'    If QPTrim$(Emp2Rec(1).EMPEACT3) <> "" Then '10/19/04
'      Emp2Rec(1).EMPEACT3 = AddDashesToGLNumber(Emp2Rec(1).EMPEACT3, Fund, Dept, Detail)
'    End If
'    TempTransRec(1).EDist(3).EAcct = Emp2Rec(1).EMPEACT3
'  End If
'  'for loop consolidates all add'l earnings and loads into
'  'TempTransRec(1).TotAdditEarn for this employee
'  For x = 1 To 3 'possibly to 5 later
'    TempTransRec(1).TotAdditEarn = TempTransRec(1).TotAdditEarn + TempTransRec(1).EAmt(x)
'  Next x
'
'  Select Case PayType$ 'this select splits up the employees
'  'wages into GL accounts based on data saved in their employee
'  'maintenance screen
'  Case "HOURLY"
'    TempTransRec(1).PayType = "H"
'    For x = 1 To 8
'      If Emp2Rec(1).EDist(x).DAmt < 0 Then
'        TempTransRec(1).TDist(x).DRHrs = 0
'      Else
'        TempTransRec(1).TDist(x).DRHrs = Emp2Rec(1).EDist(x).DAmt
'        TempTransRec(1).RegHrsWork = OldRound#(TempTransRec(1).RegHrsWork + TempTransRec(1).TDist(x).DRHrs)
'      End If
'      If QPTrim$(Emp2Rec(1).EDist(x).DAcct) <> "" Then '10/19/04
'        Emp2Rec(1).EDist(x).DAcct = AddDashesToGLNumber(Emp2Rec(1).EDist(x).DAcct, Fund, Dept, Detail)
'      End If
'      TempTransRec(1).TDist(x).DAcct = Emp2Rec(1).EDist(x).DAcct
'    Next x
'    TempTransRec(1).RegHrsPaid = TempTransRec(1).RegHrsWork
'  Case "SALARIED"
'    TotSalDist = 0
'    TempTransRec(1).PayType = "S"
'    For x = 1 To 8
'      If Emp2Rec(1).EDist(x).DAmt < 0 Then
'        TempTransRec(1).TDist(x).DPct = 0
'      Else
'        TempTransRec(1).TDist(x).DPct = Emp2Rec(1).EDist(x).DAmt
'      End If
'      TotSalDist = OldRound#(TotSalDist + TempTransRec(1).TDist(x).DPct)
'      If QPTrim$(Emp2Rec(1).EDist(x).DAcct) <> "" Then '10/19/04
'        Emp2Rec(1).EDist(x).DAcct = AddDashesToGLNumber(Emp2Rec(1).EDist(x).DAcct, Fund, Dept, Detail)
'      End If
'      TempTransRec(1).TDist(x).DAcct = Emp2Rec(1).EDist(x).DAcct
'    Next
'    'temp var = TotSalDist  TempTransRec(1).TActive = False
'    'this if statement traps for any salaried employees whose total
'    'percentage distribution does not add up to 100%
'    If TotSalDist <> 100# Then
'      TempTransRec(1).TActive = False
'      GoTo BadSalDist
'    End If
'  End Select
'
'  If PPDRec.MACTIVE = 0 Then 'added this if on 10/3/06
'    Call CalcPay(TempTransRec(1), RecNo, False)
'  End If
'
'BadSalDist:
'  Put TRHandle, RecNo, TempTransRec(1) 'TRHandle is a global
'
'End Sub
'
'Public Sub CalcPay(TransRec As TransRecType, TransRecNo As Long, ReCalcFlag As Boolean)
'  Dim DedAmts(1 To 50) As Double
'  Dim ErnAmts(1 To 3) As Double
'  Dim StateTax(1) As StateTaxRecType
'  Dim STFNHandle As Integer
'  Dim EICHandle As Integer
'  Dim EICRec As EICRecType
'
'  Dim x As Long
'  Dim LastRActiveDist As Long
'  Dim LastOActiveDist As Long
'  Dim TotalRegDist#, DistDiff#
'  Dim TotalOTDist#, LastActiveDist#
'  Dim TotalWageDist#
'  Dim PPDHandle As Integer
'  Dim PPDRec As PeriodDefaultRecType
'  Dim EmpHandle As Integer
'  Dim DHandle As Integer
'  Dim DedCodes As DedCodeRecType
'  Dim DedCnt As Integer
'  Dim TotRegDist#, DistDif#
'  'Hourly processing
'  Dim FedHandle As Integer
'  Dim FEDTAX As FederalTaxRecType
'  Dim TaxFring#, FedExempt$, StaExempt$
'  Dim PayPFreq$(1 To 7), PayFreq As Integer
'  Dim AnnualizedFedGross#
'  Dim PriorFedTax#
'  Dim ErnHandle As Integer
'  Dim ErnCodes As ErnCodeRecType
'  Dim RHandle As Integer, FedTaxDed#
'  Dim RetireRec As RetireRecType
'  Dim RCnt As Integer, EICGross#
'  Dim SocExempt$, MedExempt$
'  Dim AnnualizedStaGross#, PriorStaTax#
'  Dim WageDiff#, AR10PctDed#
'  Dim EXSDiff#, TaxableAmtFed#
'  Dim AL20PctDed#, SCStateExmpAmt#
'  Dim ExcessAmt#, Exmp1Amt#, Exmp2Amt#, Exmp3Amt#
'  Dim TotalExmp#, TaxableAmtSta#
'  Dim TblPoint As Integer
'  Dim Multi#, TCnt As Integer
'
'  Dim EmpRec3 As EmpData3Type '12/19/02
'  Dim Emp3Handle As Integer '12/19/02
'  Dim PastSSMaxWage As Boolean '12/19/2002
'  Dim FringeFlag As Boolean
'
'  FringeFlag = False
'  ErnAmts(1) = TransRec.EAmt(1)
'  ErnAmts(2) = TransRec.EAmt(2)
'  ErnAmts(3) = TransRec.EAmt(3)
'
'  OpenStateTaxFileName STFNHandle
'  Get STFNHandle, 1, StateTax(1)
'  Close STFNHandle
'
'  Select Case TransRec.PayType
'  Case "H"
'    TransRec.TotRegWage = OldRound#(TransRec.BaseRate * TransRec.RegHrsPaid)
'    TransRec.TotOTWage = OldRound#(TransRec.OTRate * TransRec.OTHrsPaid)
'    TransRec.GrossWage = OldRound(TransRec.TotRegWage + TransRec.TotOTWage)
'
'    If TransRec.GrossWage > 0 Then
'      For x = 1 To 8 'split out wage distributions
'        TransRec.TDist(x).DRWage = OldRound#(TransRec.TDist(x).DRHrs * TransRec.BaseRate)
'        TransRec.TDist(x).DOWage = OldRound#(TransRec.TDist(x).DOHrs * TransRec.OTRate)
'        If TransRec.TDist(x).DRWage > 0 Then LastRActiveDist = x
'        If TransRec.TDist(x).DOWage > 0 Then LastOActiveDist = x
'        TransRec.TDist(x).DPct = OldRound#(100 * (TransRec.TDist(x).DRWage + TransRec.TDist(x).DOWage) / TransRec.GrossWage)
'      Next x
'    Else 'deal with 0 gross wage but has an additional earning dist here
'      For x = 1 To 8 ' No gross...all from additional earnings
'        TransRec.TDist(x).DRWage = 0
'        TransRec.TDist(x).DOWage = 0
'      Next x
'      TransRec.TDist(1).DPct = 100
'    End If
'    'calculate and adjust regular wage distributions
'    If LastRActiveDist > 0 Then
'      Do
'        TotalRegDist# = 0
'        For x = 1 To 8
'          TotalRegDist# = OldRound#(TotalRegDist# + TransRec.TDist(x).DRWage)
'        Next x
'        If TotalRegDist <> TransRec.TotRegWage Then
'          If TotalRegDist# > TransRec.TotRegWage Then
'            DistDif# = OldRound#(TotalRegDist# - TransRec.TotRegWage)
'            TransRec.TDist(LastRActiveDist).DRWage = TransRec.TDist(LastRActiveDist).DRWage = DistDif#
'          ElseIf TotalRegDist# < TransRec.TotRegWage Then
'            DistDif# = OldRound#(TransRec.TotRegWage - TotalRegDist#)
'            TransRec.TDist(LastRActiveDist).DRWage = TransRec.TDist(LastRActiveDist).DRWage + DistDif#
'          End If
'        End If
'      Loop Until TotalRegDist# = TransRec.TotRegWage
'    End If
'    'calculate and adjust Overtime wage dist
'
'    Do
'      TotalOTDist# = 0
'      For x = 1 To 8
'        TotalOTDist# = OldRound#(TotalOTDist# + TransRec.TDist(x).DOWage)
'      Next x
'      If TotalOTDist# <> TransRec.TotOTWage Then
'        If TotalOTDist# > TransRec.TotOTWage Then
'          DistDif# = OldRound#(TotalOTDist# - TransRec.TotOTWage)
'          TransRec.TDist(LastOActiveDist).DOWage = TransRec.TDist(LastOActiveDist).DOWage = DistDif#
'        ElseIf TotalOTDist# < TransRec.TotOTWage Then
'          DistDif# = OldRound#(TransRec.TotOTWage - TotalOTDist#)
'          TransRec.TDist(LastOActiveDist).DOWage = TransRec.TDist(LastOActiveDist).DOWage + DistDif#
'        End If
'      End If
'    Loop Until TotalOTDist# = TransRec.TotOTWage
'  Case "S"
'    If TransRec.PaySFlag <> "Y" And TransRec.PaySFlag <> "N" Then 'BB change
'      TransRec.PaySFlag = "Y"
'    End If
'
'    If TransRec.PaySFlag = "Y" Then
'      If Not ReCalcFlag Then
'        TransRec.GrossWage = OldRound#(TransRec.BaseRate)
'        TransRec.TotRegWage = OldRound#(TransRec.BaseRate)
'      Else
'        TransRec.GrossWage = TransRec.TotRegWage
'      End If
'
'      For x = 1 To 8
'        TransRec.TDist(x).DRWage = OldRound#((TransRec.TDist(x).DPct * 0.01) * TransRec.GrossWage)
'        If TransRec.TDist(x).DRWage > 0 Then LastActiveDist = x
'      Next
'
'      If LastActiveDist = 0 Then
'        TransRec.TActive = False
'        Exit Sub
'      End If
'
'      Do
'        TotalWageDist# = 0
'        For x = 1 To 8
'          TotalWageDist# = OldRound#(TotalWageDist# + TransRec.TDist(x).DRWage)
'        Next
'
'        If TotalWageDist# <> TransRec.GrossWage Then
'          If TotalWageDist# > TransRec.GrossWage Then
'            DistDif# = OldRound#(TotalWageDist# - TransRec.GrossWage)
'            TransRec.TDist(LastActiveDist).DRWage = TransRec.TDist(LastActiveDist).DRWage - DistDif#
'          ElseIf TotalWageDist# < TransRec.GrossWage Then
'            DistDif# = OldRound#(TransRec.GrossWage - TotalWageDist#)
'            TransRec.TDist(LastActiveDist).DRWage = TransRec.TDist(LastActiveDist).DRWage + DistDif#
'          End If
'        End If
'      Loop Until TotalWageDist# = TransRec.GrossWage
'
'    Else
'      'deal with zero gross wage, but has an additional earning dist here
'      For x = 1 To 8
'        TransRec.TDist(x).DRWage = 0
'        TransRec.TDist(x).DPct = 0
'      Next x
'
'      TransRec.TDist(1).DPct = 100
'      TransRec.GrossWage = 0
'      TransRec.TotRegWage = 0
'      TransRec.BaseRate = 0
'    End If
'  End Select
'  '*****************************************************************
'  TransRec.GrossPay = OldRound#(TransRec.GrossWage + TransRec.TotAdditEarn)
'  TransRec.FedGrossPay = TransRec.GrossPay
'  TransRec.StaGrossPay = TransRec.GrossPay
'  TransRec.SocGrossPay = TransRec.GrossPay
'  TransRec.MedGrossPay = TransRec.GrossPay
'  TransRec.RetGrossPay = TransRec.GrossPay
'
'  '******************************************************************
'
'  'Deduction processing
'  OpenPPDefaultFile PPDHandle
'  Get PPDHandle, 1, PPDRec
'  Close PPDHandle
'
'  OpenEmpData2File EmpHandle
'  Get EmpHandle, TransRecNo, Emp2Rec(1)
'  Close EmpHandle
'  If Not ReCalcFlag Then 'this employee is not PR active yet
'  'so build his transaction file with data selected in
'  'payroll defaults
'    TransRec.TotDedAmt = 0
'    For x = 1 To 50
'      If PPDRec.UseDed(x) = "Y" Then
'        If Emp2Rec(1).EmpDed(x).DAmt > 0 Then
'          Select Case Mid(Emp2Rec(1).EmpDed(x).DPct, 1, 1)
'          Case "A"  'Amount
'            DedAmts#(x) = Emp2Rec(1).EmpDed(x).DAmt
'          Case "P" ' Percent
'            Select Case Emp2Rec(1).EmpDed(x).DOTI
'            Case "N"  'No
'              DedAmts#(x) = OldRound#(TransRec.TotRegWage * (Emp2Rec(1).EmpDed(x).DAmt * 0.01))
'            Case Else '" ", "Y", "" 'yes or space (blank)
'              DedAmts#(x) = OldRound#(TransRec.GrossPay * (Emp2Rec(1).EmpDed(x).DAmt * 0.01))
'            End Select
'          End Select
'          TransRec.DAmt(x) = DedAmts#(x)
'        Else
'          TransRec.DAmt(x) = 0
'        End If
'      End If
'      TransRec.TotDedAmt = OldRound#(TransRec.TotDedAmt + TransRec.DAmt(x))
'    Next
'  Else  'recalc flag is true so his transaction data has
'  'already been PR default updated so load screen as the
'  'data has been stored
'    TransRec.TotDedAmt = 0
'    For x = 1 To 50
'      If TransRec.DAmt(x) > 0 Then
'        DedAmts#(x) = TransRec.DAmt(x)
'        Select Case Mid(Emp2Rec(1).EmpDed(x).DPct, 1, 1)
'        Case "A" 'Amount
'          DedAmts#(x) = TransRec.DAmt(x)
'        Case "P" 'Percent
'          Select Case Emp2Rec(1).EmpDed(x).DOTI
'          Case "N"  'no
'            DedAmts#(x) = OldRound#(TransRec.TotRegWage * (Emp2Rec(1).EmpDed(x).DAmt * 0.01))
'          Case "", "Y"
'            DedAmts#(x) = OldRound#(TransRec.GrossPay * (Emp2Rec(1).EmpDed(x).DAmt * 0.01))
'          End Select
'        End Select
'      End If
'      TransRec.DAmt(x) = DedAmts#(x)
'      TransRec.TotDedAmt = OldRound#(TransRec.TotDedAmt + TransRec.DAmt(x))
'    Next x
'  End If
'
'  '**************************************************************
'  'calculations taxable amounts "DEDUCTIONS"
'
'  OpenDedCodeFile DHandle
'  DedCnt = LOF(DHandle) / Len(DedCodes)
'  For x = 1 To DedCnt
'    Get DHandle, x, DedCodes
'    If DedCodes.DCFWT1 = "Y" Then
'      If DedAmts#(x) > 0 Then
'        TransRec.FedGrossPay = OldRound#(TransRec.FedGrossPay - DedAmts#(x))
'      End If
'    End If
'
'    If DedCodes.DCSWT1 = "Y" Then
'      If DedAmts#(x) > 0 Then
'        TransRec.StaGrossPay = OldRound#(TransRec.StaGrossPay - DedAmts#(x))
'      End If
'    End If
'
'    If DedCodes.DCSOC1 = "Y" Then
'      If DedAmts#(x) > 0 Then
'        TransRec.SocGrossPay = OldRound#(TransRec.SocGrossPay - DedAmts#(x))
'      End If
'    End If
'
'    If DedCodes.DCMED1 = "Y" Then
'      If DedAmts#(x) > 0 Then
'        TransRec.MedGrossPay = OldRound#(TransRec.MedGrossPay - DedAmts#(x))
'      End If
'    End If
'  Next x
'  Close DHandle
'  'if retamt > 0 and not tax exempt then adjust fedgrosspay
'  'calc taxable amounts "EARNINGS"
'
'  OpenErnCodeFile ErnHandle
'  For x = 1 To 3
'    Get ErnHandle, x, ErnCodes
'    If ErnCodes.ERNFWT1 = "N" Then
'      TransRec.FedGrossPay = OldRound#(TransRec.FedGrossPay - ErnAmts#(x))
'    End If
'    If ErnCodes.ERNSWT1 = "N" Then
'      TransRec.StaGrossPay = OldRound#(TransRec.StaGrossPay - ErnAmts#(x))
'    End If
'    If ErnCodes.ERNSOC1 = "N" Then
'      TransRec.SocGrossPay = OldRound#(TransRec.SocGrossPay - ErnAmts#(x))
'    End If
'    If ErnCodes.ERNMED1 = "N" Then
'      TransRec.MedGrossPay = OldRound#(TransRec.MedGrossPay - ErnAmts#(x))
'    End If
'    If ErnCodes.ERNRET1 = "N" Then
'      TransRec.RetGrossPay = OldRound#(TransRec.RetGrossPay - ErnAmts#(x))
'    End If
'  Next
'  TransRec.Less401k(1) = False
'  TransRec.Less401k(2) = False
'  TransRec.Less401k(3) = False
'  For x = 1 To 3 'added 10/08/03
'    Get ErnHandle, x, ErnCodes
'    If ErnCodes.EarnYN = "N" Then
'      TransRec.Less401k(x) = True
'    End If
'  Next x
'
'  Close ErnHandle
'  'get soc and gross here
'  'calculate retirement W/H
'
'  OpenRetFile RHandle
'  RCnt = LOF(RHandle) / Len(RetireRec)
''  If Len(Emp2Rec(1).EMPRETNO) > 0 Then  'changed by BB
''  If Len(Emp2Rec(1).EMPRETNO) > 0 And Mid(Emp2Rec(1).EMPRETNO, 1, 1) <> "T" And Mid(Emp2Rec(1).EMPRETNO, 1, 1) <> "R" Then '7/22/2004
'  If Len(QPTrim$(Emp2Rec(1).EMPRETNO)) > 0 And Mid(Emp2Rec(1).EMPRETNO, 1, 1) <> "T" And Mid(Emp2Rec(1).EMPRETNO, 1, 1) <> "R" Then '8/23/2005
'    For x = 1 To 6
'      Get RHandle, x, RetireRec
'      If UCase$(QPTrim$(RetireRec.TYPEDES1)) = UCase$(QPTrim$(Emp2Rec(1).EMPRETTP)) Then
'        Exit For
'      End If
'    Next x
'
'    Close RHandle
'
'    Select Case x
'    Case 1 To 6
'      Select Case RetireRec.TYPEOT1
'      Case "Y"  'include overtime in retirement calc!
'        TransRec.RetireAmt = OldRound#(TransRec.RetGrossPay * (RetireRec.TYPEWH1 * 0.01))
'        TransRec.MatchRetAmt = OldRound#(TransRec.RetGrossPay * (RetireRec.TYPEM1 * 0.01))
'      Case "N" 'nope don't include OT
'        TransRec.RetGrossPay = OldRound#(TransRec.RetGrossPay - TransRec.TotOTWage)
'        TransRec.RetireAmt = OldRound#(TransRec.RetGrossPay * (RetireRec.TYPEWH1 * 0.01))
'        TransRec.MatchRetAmt = OldRound#(TransRec.RetGrossPay * (RetireRec.TYPEM1 * 0.01))
'      End Select
'    Case Else
'      TransRec.RetireAmt = 0
'      TransRec.MatchRetAmt = 0
'    End Select
'  Else
'    TransRec.RetireAmt = 0
'    TransRec.MatchRetAmt = 0
'  End If
'  Close RHandle
'  'adjust taxable amounts after retirement calculations
'  If x < 7 Then 'x will only be > 6 if it ran all the
'  'way thru the last for loop without finding a match
'    If QPTrim$(RetireRec.TYPETD1) = "Y" Then
'      TransRec.FedGrossPay = OldRound#(TransRec.FedGrossPay - TransRec.RetireAmt)
'      TransRec.StaGrossPay = OldRound#(TransRec.StaGrossPay - TransRec.RetireAmt)
'    End If
'  End If
'
'  'add retirement to total deductions
'  TransRec.TotDedAmt = OldRound#(TransRec.TotDedAmt + TransRec.RetireAmt)
'  SocExempt$ = QPTrim$(Emp2Rec(1).EMPSOCX)
'  If SocExempt$ = "" Then SocExempt = "N" 'makes DOS transfers work
'  MedExempt$ = QPTrim$(Emp2Rec(1).EMPMEDX)
'  If MedExempt$ = "" Then MedExempt$ = "N"
'  'calculate social security W/H
'  OpenFedTaxFile FedHandle
'  Get FedHandle, 1, FEDTAX
'  Close FedHandle
'
'  '>>>>>>>>>>>>>>>12/19/2002
'  OpenEmpData3File Emp3Handle
'  Get Emp3Handle, TransRecNo, EmpRec3
'  Close Emp3Handle
'  PastSSMaxWage = False
'  Select Case SocExempt$
'  Case "N"
'    TransRec.SocTaxAmt = OldRound#(TransRec.SocGrossPay * (FEDTAX.FTMEMPSS * 0.01)) 'figure this pay period's SS withholding
'    If TransRec.SocGrossPay + EmpRec3.YTDSocGrossPay > FEDTAX.FTMSSMW Then 'if Social Gross Pay is past the maximum
'      'next line added on 11/6/03 to accommodate the reduction of social gross wage
'      '(in addition to social tax) for those making more than the maximum social security gross wage
'      TransRec.SocGrossPay = FEDTAX.FTMSSMW - EmpRec3.YTDSocGrossPay
'      If TransRec.SocGrossPay < 0 Then TransRec.SocGrossPay = 0
'      If EmpRec3.YTDSocial + TransRec.SocTaxAmt > OldRound(FEDTAX.FTMSSMW * (FEDTAX.FTMEMPSS * 0.01)) Then 'now compare the YTD total plus
'      'this pay period's total against the maximum SS withholding allowable and if it is too much then adjust accordingly
'        TransRec.SocTaxAmt = OldRound((FEDTAX.FTMSSMW * FEDTAX.FTMEMPSS * 0.01) - EmpRec3.YTDSocial)
'        PastSSMaxWage = True
'      End If
'    End If
'
' '>>>>>>>>>>>>>>>>>>>>12/19/02
'    If PastSSMaxWage = True Then GoTo Max1 'if max wage is reached then no tax fringe tax is paid '1/2/03
'    If TransRec.TaxFring > 0 Then
'      TaxFring# = OldRound#((TransRec.TaxFring * FEDTAX.FTMEMPSS) * 0.01)
'      TransRec.SocTaxAmt = OldRound#(TransRec.SocTaxAmt + TaxFring#)
'    End If
'
'Max1:
'    If PastSSMaxWage = True Then
'      TransRec.MatchSocAmt = OldRound#((FEDTAX.FTMSSMW * FEDTAX.FTMEMRSS * 0.01) - EmpRec3.YTDSocial)
'    Else
'      TransRec.MatchSocAmt = OldRound#((TransRec.SocGrossPay * FEDTAX.FTMEMRSS) * 0.01)
'    End If
'
'    If PastSSMaxWage = True Then
'      PastSSMaxWage = False
'      GoTo Max2 'if max wage is reached then no tax fringe tax is paid '1/2/03
'    End If
'    If TransRec.TaxFring > 0 Then
'      TaxFring# = OldRound#((TransRec.TaxFring * FEDTAX.FTMEMRSS) * 0.01)
'      TransRec.MatchSocAmt = OldRound#(TransRec.MatchSocAmt + TaxFring#)
'    End If
'Max2:
'  Case "Y"
'    TransRec.SocTaxAmt = 0
'    TransRec.SocGrossPay = 0
'    TransRec.MatchSocAmt = 0
'  End Select
'
'  'calculations medicare W/H
'  Select Case MedExempt$
'  Case "N"
'    TransRec.MedTaxAmt = OldRound#((TransRec.MedGrossPay * FEDTAX.FTMEMPM) * 0.01)
'    If TransRec.TaxFring > 0 Then
'      TaxFring# = OldRound#((TransRec.TaxFring * FEDTAX.FTMEMPM) * 0.01)
'      TransRec.MedTaxAmt = OldRound#(TransRec.MedTaxAmt + TaxFring#)
'    End If
'    TransRec.MatchMedAmt = OldRound#((TransRec.MedGrossPay * FEDTAX.FTMEMPM) * 0.01)
'    If TransRec.TaxFring > 0 Then
'      TaxFring# = OldRound#((TransRec.TaxFring * FEDTAX.FTMEMRM) * 0.01)
'      TransRec.MatchMedAmt = OldRound#(TransRec.MatchMedAmt + TaxFring#)
'    End If
'
'  Case "Y"
'    TransRec.MedTaxAmt = 0
'    TransRec.MedGrossPay = 0
'    TransRec.MatchMedAmt = 0
'  End Select
'
'  '-------------------------------------------
'  'Start of State and Federal tax calculations
'
'  FedExempt$ = QPTrim$(Emp2Rec(1).EMPFEDX)
'  If FedExempt$ = "" Then FedExempt = "N" 'makes DOS data transfers work
'  StaExempt$ = QPTrim$(Emp2Rec(1).EMPSTAX)
'  If StaExempt$ = "" Then StaExempt$ = "N"
'
'  PayPFreq$(1) = UCase$("Weekly")
'  PayPFreq$(2) = UCase$("Bi-weekly")
'  PayPFreq$(3) = UCase$("Semi-Monthly")
'  PayPFreq$(4) = UCase$("Monthly")
'  PayPFreq$(5) = UCase$("Quarterly")
'  PayPFreq$(6) = UCase$("Semi-Annually")
'  PayPFreq$(7) = UCase$("Annually")
'
'  For x = 1 To 7
'    If UCase$(QPTrim$(Emp2Rec(1).EMPPFREQ)) = UCase$(PayPFreq$(x)) Then
'      Exit For
'    End If
'  Next
'
'  Select Case x
'  Case 1
'    PayFreq = 52
'  Case 2
'    PayFreq = 26
'  Case 3
'    PayFreq = 24
'  Case 4
'    PayFreq = 12
'  Case 5
'    PayFreq = 4
'  Case 6
'    PayFreq = 2
'  Case 7
'    PayFreq = 1
'  End Select
'  Select Case FedExempt$  'are they fed exempt
'  Case "N", ""  'no they aren't
'    Select Case QPTrim$(Emp2Rec(1).EMPFEDO2) 'using fixed amount or percent
'    Case "" ' no
'      AnnualizedFedGross# = OldRound#(TransRec.FedGrossPay * PayFreq)
'      GoSub CalcFedTax
'      If TransRec.TaxFring > 0 Then
'        FringeFlag = True 'added 5/14/03
'        PriorFedTax# = TransRec.FedTaxAmt
'        'this variable has to remain this name
'        'because it is used in the gosub to calcfedtax
'        AnnualizedFedGross# = TransRec.TaxFring
'        GoSub CalcFedTax
'        FringeFlag = False 'added 5/14/03
'        TransRec.FedTaxAmt = OldRound#(TransRec.FedTaxAmt + PriorFedTax#)
'      End If
'    Case "P"  'using fixed percent
'      TransRec.FedTaxAmt = OldRound#(TransRec.FedGrossPay * (Emp2Rec(1).EMPFEDO1 * 0.01))
'    Case "A"  'using fixed amount
'      TransRec.FedTaxAmt = Emp2Rec(1).EMPFEDO1
'    End Select
'  Case "Y"  'yes, they are fed exempt
'    TransRec.FedTaxAmt = 0
'    AnnualizedFedGross# = 0
'  End Select
'
'  Select Case StaExempt$
'  Case "N", "" 'no they aren't
'    Select Case QPTrim$(Emp2Rec(1).EMPSTAO2) 'using fixed amount or percent?
'    Case ""  'no
'      AnnualizedStaGross# = OldRound#(TransRec.StaGrossPay * PayFreq)
'      GoSub CalcStaTax
'      If TransRec.TaxFring > 0 Then
'        FringeFlag = True 'added 5/14/03
'        PriorStaTax# = TransRec.StaTaxAmt
'        'this variable has to remain this name
'        'because it is used in the gosub to CalcStaTax
'        AnnualizedStaGross# = TransRec.TaxFring
'        GoSub CalcStaTax
'        FringeFlag = False 'added 5/14/03
'        TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + PriorStaTax#)
'      End If
'    Case "P"  'using fixed percent 'added CInt on 6/7/06
'      TransRec.StaTaxAmt = CInt(OldRound#(TransRec.StaGrossPay * (Emp2Rec(1).EMPSTAO1 * 0.01)))
'    Case "A"  'using fixed amount
'      TransRec.StaTaxAmt = CInt(Emp2Rec(1).EMPSTAO1)
'    End Select
'  Case "Y"
'    TransRec.StaTaxAmt = 0
'    AnnualizedStaGross# = 0
'  End Select
'
'  TransRec.TotTaxAmt = OldRound#(TransRec.StaTaxAmt + TransRec.FedTaxAmt)
'  TransRec.TotTaxAmt = OldRound#(TransRec.TotTaxAmt + TransRec.MedTaxAmt + TransRec.SocTaxAmt)
'  TransRec.NetPay = OldRound#((TransRec.GrossPay - TransRec.TotTaxAmt) - TransRec.TotDedAmt)
'
'  If TransRec.GrossPay <= 0 Then
'    TransRec.TActive = False
'  End If
'  EICGross# = OldRound#(TransRec.FedGrossPay * PayFreq)
'  OpenEICFile EICHandle
'  Get EICHandle, 1, EICRec
'  Close EICHandle
'
'  Select Case QPTrim$(Emp2Rec(1).EMPEIC)
'  Case "0", ""
'    TransRec.EICAmt = 0
'  Case "1"
'    Select Case EICGross#
'    Case Is < EICRec.EIC(1).EIC1NVR0
'      TransRec.EICAmt = OldRound#((TransRec.FedGrossPay * EICRec.EIC(1).EIC1AMT0) * 0.01)
'    Case EICRec.EIC(1).EIC1OVR1 + 1 To EICRec.EIC(1).EIC1NVR1
'      TransRec.EICAmt = OldRound#(EICRec.EIC(1).EIC1AMT1 / PayFreq)
'    Case Is > EICRec.EIC(1).EIC1OVR2
'      WageDiff# = OldRound#(EICGross - EICRec.EIC(1).EIC1EXES)
'      EXSDiff# = OldRound#((WageDiff# * EICRec.EIC(1).EIC1LESS) * 0.01)
'      TransRec.EICAmt = OldRound#((EICRec.EIC(1).EIC1AMT1 - EXSDiff#) / PayFreq)
'      If TransRec.EICAmt < 1 Then
'        TransRec.EICAmt = 0
'      End If
'    End Select
'  Case "2"
'    Select Case EICGross#
'    Case Is < EICRec.EIC(2).EIC1NVR0
'      TransRec.EICAmt = OldRound#((TransRec.FedGrossPay * EICRec.EIC(2).EIC1AMT0) * 0.01)
'    Case EICRec.EIC(2).EIC1OVR1 + 1 To EICRec.EIC(2).EIC1NVR1
'      TransRec.EICAmt = OldRound#(EICRec.EIC(2).EIC1AMT1 / PayFreq)
'    Case Is > EICRec.EIC(2).EIC1OVR2
'      WageDiff# = OldRound#(EICGross# - EICRec.EIC(2).EIC1EXES)
'      EXSDiff# = OldRound#((WageDiff# * EICRec.EIC(2).EIC1LESS) * 0.01)
'      TransRec.EICAmt = OldRound#((EICRec.EIC(2).EIC1AMT1 - EXSDiff#) / PayFreq)
'      If TransRec.EICAmt < 1 Then
'        TransRec.EICAmt = 0
'      End If
'    End Select
'  End Select
'  TransRec.NetPay = OldRound(TransRec.EICAmt + TransRec.NetPay)
'  If TransRec.FedGrossPay < 0 Then TransRec.FedGrossPay = 0
'  If TransRec.StaGrossPay < 0 Then TransRec.StaGrossPay = 0
'  If TransRec.SocGrossPay < 0 Then TransRec.SocGrossPay = 0
'  If TransRec.MedGrossPay < 0 Then TransRec.MedGrossPay = 0
'  If TransRec.RetGrossPay < 0 Then TransRec.RetGrossPay = 0
'
'  Exit Sub
'
'CalcFedTax:
'  Select Case QPTrim$(Emp2Rec(1).EMPFEDS)  'single or married?
'  Case "M"
'    If Emp2Rec(1).EMPFEDA < 0 Then Emp2Rec(1).EMPFEDA = 0
'    TaxableAmtFed# = OldRound#(AnnualizedFedGross# - (Emp2Rec(1).EMPFEDA * FEDTAX.FTMSDAA))
'    If TaxableAmtFed# < FEDTAX.FTM(3, 1) Then
'      TransRec.FedTaxAmt = 0
'    Else
'      For x = 1 To 10
'        If FEDTAX.FTM(3, x) > TaxableAmtFed# Then Exit For
'      Next
'      x = x - 1
'      TaxableAmtFed# = OldRound#(TaxableAmtFed# - FEDTAX.FTM(3, x))
'      TransRec.FedTaxAmt = OldRound#(OldRound#(FEDTAX.FTM(1, x) + (TaxableAmtFed# * (FEDTAX.FTM(2, x) * 0.01))) / PayFreq)
'    End If
'  Case "S", ""
'    If Emp2Rec(1).EMPFEDA < 0 Then Emp2Rec(1).EMPFEDA = 0
'    TaxableAmtFed# = OldRound#(AnnualizedFedGross# - (Emp2Rec(1).EMPFEDA * FEDTAX.FTSSDAA))
'    If TaxableAmtFed# < FEDTAX.FTS(3, 1) Then
'      TransRec.FedTaxAmt = 0
'    Else
'      For x = 1 To 10
'        If FEDTAX.FTS(3, x) > TaxableAmtFed# Then Exit For
'      Next
'      x = x - 1
'      TaxableAmtFed# = OldRound#(TaxableAmtFed# - FEDTAX.FTS(3, x))
'      TransRec.FedTaxAmt = OldRound#(OldRound#(FEDTAX.FTS(1, x) + (TaxableAmtFed# * (FEDTAX.FTS(2, x) * 0.01))) / PayFreq)
'    End If
'  End Select
''  If Emp2Rec(1).EMPFEDAA > 0 Then
'  If Emp2Rec(1).EMPFEDAA > 0 And FringeFlag = False Then 'added 5/14/03
'    TransRec.FedTaxAmt = OldRound#(TransRec.FedTaxAmt + Emp2Rec(1).EMPFEDAA)
'  End If
'TaxFring:
'  Return
'
'CalcStaTax:
'
'  Select Case TaxText(1)
'  Case "OK"
'  'add some kind of an include section here for various state calc's
'  'look at employer's state in the controll file to determin which
'  'state tax tables to use or edit
'  '  STOP
'  If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'  Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
'    Case Left$(TaxText$(2), 1), ""
'      If AnnualizedStaGross# > StateTax(1).TAX101 Then
'        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX101
'        AnnualizedStaGross# = StateTax(1).TAX101
'      End If
'
'      Exmp1Amt# = OldRound(Emp2Rec(1).EMPSTAA * StateTax(1).TAX102)
'                    'number of allounces * std ded amt
'      Exmp2Amt# = OldRound(AnnualizedStaGross# * (StateTax(1).TAX105 * 0.01))
'                    'Annual gross wages * Std Ded Rate%
'
'      Select Case Exmp2Amt#
'      Case Is >= StateTax(1).TAX107
'        'if exempt 2 amt is greater than max std amt
'        Exmp2Amt# = StateTax(1).TAX107
'      Case Is < StateTax(1).TAX106
'        'if exempt 2 amt is less than min std amt
'        Exmp2Amt# = StateTax(1).TAX106
'      End Select
'
'      Exmp3Amt# = OldRound((AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX103) - StateTax(1).TAX104))
'
'      Exmp3Amt# = OldRound(Exmp3Amt# * (StateTax(1).TAX105 * 0.01))
'
'      If Exmp3Amt# < 0 Then Exmp3Amt# = 0
'      TotalExmp# = OldRound(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - TotalExmp#)
'
'      If TaxableAmtSta# < 0 Then
'        TaxableAmtSta# = 0
'      End If
'      For TCnt = 1 To 12
'        If TaxableAmtSta# < StateTax(1).STS(3, TCnt) Then
'          Exit For
'        End If
'      Next
'      TblPoint = TCnt
'
'      If TblPoint = 1 Then
'        TransRec.StaTaxAmt = OldRound#(TaxableAmtSta# * (StateTax(1).STS(2, TblPoint) * 0.01) / PayFreq)
'      Else
'        TransRec.StaTaxAmt = OldRound#((StateTax(1).STS(1, TCnt) + (TaxableAmtSta# - StateTax(1).STS(3, TCnt - 1)) * (StateTax(1).STS(2, TCnt) * 0.01)) / PayFreq)
'      End If
'
'      If ExcessAmt# > 0 Then
'        For RCnt = 12 To 1 Step -1
'          If StateTax(1).STS(2, RCnt) > 0 Then
'            Exit For
'          End If
'        Next
'        Multi# = OldRound#(StateTax(1).STS(2, RCnt) * 0.01)
'        TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
'      End If
'
'    Case Left$(TaxText$(3), 1) 'changed by BB
'      If AnnualizedStaGross# > StateTax(1).TAX201 Then
'        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX201
'        AnnualizedStaGross# = StateTax(1).TAX201
'      End If
'
'      Exmp1Amt# = OldRound(Emp2Rec(1).EMPSTAA * StateTax(1).TAX202)
'                    'number of allounces * std ded amt
'      Exmp2Amt# = OldRound(AnnualizedStaGross# * (StateTax(1).TAX205 * 0.01))
'                    'Annual gross wages * Std Ded Rate%
'
'      Select Case Exmp2Amt#
'      Case Is >= StateTax(1).TAX207
'        'if exempt 2 amt is greater than max std amt
'        Exmp2Amt# = StateTax(1).TAX207
'      Case Is < StateTax(1).TAX206
'        'if exempt 2 amt is less than min std amt
'        Exmp2Amt# = StateTax(1).TAX206
'      End Select
'
'      Exmp3Amt# = OldRound((AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX203) - StateTax(1).TAX204))
'
'      Exmp3Amt# = OldRound(Exmp3Amt# * (StateTax(1).TAX205 * 0.01))
'
'      If Exmp3Amt# < 0 Then Exmp3Amt# = 0
'
'      TotalExmp# = OldRound(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - TotalExmp#)
'
'      If TaxableAmtSta# < 0 Then
'        TaxableAmtSta# = 0
'      End If
'
'      For TCnt = 1 To 12
'        If TaxableAmtSta# < StateTax(1).STM(3, TCnt) Then
'          Exit For
'        End If
'      Next
'
'      TblPoint = TCnt
'
'      If TblPoint = 1 Then
'        TransRec.StaTaxAmt = OldRound#(TaxableAmtSta# * (StateTax(1).STM(2, TblPoint) * 0.01) / PayFreq)
'      Else
'        TransRec.StaTaxAmt = OldRound#((StateTax(1).STM(1, TCnt) + (TaxableAmtSta# - StateTax(1).STM(3, TCnt - 1)) * (StateTax(1).STM(2, TCnt) * 0.01)) / PayFreq)
'      End If
'
'      If ExcessAmt# > 0 Then
'        For RCnt = 12 To 1 Step -1
'          If StateTax(1).STM(2, RCnt) > 0 Then
'            Exit For
'          End If
'        Next
'        Multi# = OldRound#(StateTax(1).STM(2, RCnt) * 0.01)
'        TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
'      End If
'
'    Case Left$(TaxText$(4), 1)
'
'      If AnnualizedStaGross# > StateTax(1).TAX301 Then
'        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX301
'        AnnualizedStaGross# = StateTax(1).TAX301
'      End If
'
'      Exmp1Amt# = OldRound(Emp2Rec(1).EMPSTAA * StateTax(1).TAX302)
'                    'number of allounces * std ded amt
'      Exmp2Amt# = OldRound(AnnualizedStaGross# * (StateTax(1).TAX305 * 0.01))
'                    'Annual gross wages * Std Ded Rate%
'
'      Select Case Exmp2Amt#
'      Case Is >= StateTax(1).TAX307
'        'if exempt 2 amt is greater than max std amt
'        Exmp2Amt# = StateTax(1).TAX307
'      Case Is < StateTax(1).TAX306
'        'if exempt 2 amt is less than min std amt
'        Exmp2Amt# = StateTax(1).TAX306
'      End Select
'
'      Exmp3Amt# = OldRound((AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX303) - StateTax(1).TAX304))
'
'      Exmp3Amt# = OldRound(Exmp3Amt# * (StateTax(1).TAX305 * 0.01))
'
'      If Exmp3Amt# < 0 Then Exmp3Amt# = 0
'
'      TotalExmp# = OldRound(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - TotalExmp#)
'
'      If TaxableAmtSta# < 0 Then
'        TaxableAmtSta# = 0
'      End If
'
'      For TCnt = 1 To 12
'        If TaxableAmtSta# < StateTax(1).STH(3, TCnt) Then
'          Exit For
'        End If
'      Next
'
'      TblPoint = TCnt
'
'      If TblPoint = 1 Then
'        TransRec.StaTaxAmt = OldRound#(TaxableAmtSta# * (StateTax(1).STH(2, TblPoint) * 0.01) / PayFreq)
'      Else
'        TransRec.StaTaxAmt = OldRound#((StateTax(1).STH(1, TCnt) + (TaxableAmtSta# - StateTax(1).STH(3, TCnt - 1)) * (StateTax(1).STH(2, TCnt) * 0.01)) / PayFreq)
'      End If
'
'
'      If ExcessAmt# > 0 Then
'        For RCnt = 12 To 1 Step -1
'          If StateTax(1).STH(2, RCnt) > 0 Then
'            Exit For
'          End If
'        Next
'        Multi# = OldRound#(StateTax(1).STH(2, RCnt) * 0.01)
'        TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
'      End If
'
'  End Select
'
'
'  If Emp2Rec(1).EMPSTAAA > 0 And PriorStaTax# = 0 Then
'    TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'  End If
'  TransRec.StaTaxAmt = CInt(TransRec.StaTaxAmt) 'added 01/2/07
'
'
'  Case "NC", "VA"
'    'add some kind of an include section here for various state calc's
''look at employer's state in the controll file to determin which
''state tax tables to use or edit
'
'    Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
'      Case "S", ""
'        If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'        TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX106))
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX103)
'        If TaxableAmtSta# <= 0 Then
'          TransRec.StaTaxAmt = 0
'        ElseIf TaxableAmtSta# < StateTax(1).STS(3, 1) Then
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# < StateTax(1).STS(3, 2) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq)
'        Else
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq)
'        End If
'
'      Case "M"
'        If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'        TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX206))
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX203)
'        If TaxableAmtSta# <= 0 Then
'          TransRec.StaTaxAmt = 0
'        ElseIf TaxableAmtSta# < StateTax(1).STM(3, 1) Then
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STM(3, 1) And TaxableAmtSta# < StateTax(1).STM(3, 2) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 1))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STM(3, 2) And TaxableAmtSta# < StateTax(1).STM(3, 3) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 2))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * 0.01))) / PayFreq)
'        Else
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 3))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * 0.01))) / PayFreq)
'        End If
'
'      Case "H"
'        TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX306))
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX303)
'        If TaxableAmtSta# <= 0 Then
'          TransRec.StaTaxAmt = 0
'        ElseIf TaxableAmtSta# < StateTax(1).STH(3, 1) Then
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STH(3, 1) And TaxableAmtSta# < StateTax(1).STH(3, 2) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 1))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * 0.01))) / PayFreq)
'        ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq)
'        Else
'          TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'          TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq)
'        End If
'    End Select
'
''    If Emp2Rec(1).EMPSTAAA > 0 Then
'    If Emp2Rec(1).EMPSTAAA > 0 And FringeFlag = False Then 'added 5/14/03
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'    End If
'
'    If TaxText(1) = "NC" Or TaxText(1) = "VA" Then 'added VA on 01/2/07
'      TransRec.StaTaxAmt = CInt(TransRec.StaTaxAmt) '3/28/03 Rounds individual state withholding
'      'to the nearest whole dollar to comply with NC state law
'    End If
'
'  Case "SC"
'    If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'    TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX106))
'
'    If Emp2Rec(1).EMPSTAA > 0 Then
'      SCStateExmpAmt# = OldRound#(AnnualizedStaGross# * 0.1)
'      If SCStateExmpAmt# > StateTax(1).TAX103 Then
'        SCStateExmpAmt# = StateTax(1).TAX103
'      End If
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - SCStateExmpAmt#)
'    End If
'
'    If TaxableAmtSta# <= 0 Then
'      TransRec.StaTaxAmt = 0
'    ElseIf TaxableAmtSta# < StateTax(1).STS(3, 1) Then  '1
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))) / PayFreq)
'    ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# < StateTax(1).STS(3, 2) Then '2
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))) / PayFreq) '3
'    ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq) '4
'
''******
'    ElseIf TaxableAmtSta# > StateTax(1).STS(3, 3) And TaxableAmtSta# < StateTax(1).STS(3, 4) Then
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq) '5
'
'    ElseIf TaxableAmtSta# > StateTax(1).STS(3, 4) And TaxableAmtSta# < StateTax(1).STS(3, 5) Then
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 4))
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * 0.01))) / PayFreq) '6
''******
'    Else
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 5))
'      TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * 0.01))) / PayFreq)
'    End If
'
'  If Emp2Rec(1).EMPSTAAA > 0 Then
'    TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'  End If
'
'
'  Case "GA"
'  '--Georgia state tax calculation
'
'  TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
'  If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'  Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
'
'    Case "F"  'Table F - Married one income, "H",
'
'      '--Get Standard Deduction
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX101)
'
'      '--Get personal and dependendant deductions if claiming at least one allowance
'      If Emp2Rec(1).EMPSTAA > 0 Then
'        '--Personal Allowance Amount
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX102)
'        '--Dependent Exemption
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX103))
'      End If
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# <= StateTax(1).STS(3, 2) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# <= StateTax(1).STS(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# <= StateTax(1).STS(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 3) And TaxableAmtSta# <= StateTax(1).STS(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 4) And TaxableAmtSta# <= StateTax(1).STS(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 5) And TaxableAmtSta# <= StateTax(1).STS(3, 6) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 6))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * 0.01))) / PayFreq)
'      End If
'
'    Case "G" 'Table G ,"M", - Married Filing Joint (dual income)"
'      '--Get Standard Deduction
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX201)
'
'      '--Get personal and dependendant deductions if claiming at least one allowance
'      If Emp2Rec(1).EMPSTAA > 0 Then
'        '--Personal Allowance Amount
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX202)
'        '--Dependent Exemption
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX203))
'      End If
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# <= StateTax(1).STM(3, 2) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 1) And TaxableAmtSta# <= StateTax(1).STM(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 2) And TaxableAmtSta# <= StateTax(1).STM(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 3) And TaxableAmtSta# <= StateTax(1).STM(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 4) And TaxableAmtSta# <= StateTax(1).STM(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 5) And TaxableAmtSta# <= StateTax(1).STM(3, 6) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 5) + (TaxableAmtSta# * (StateTax(1).STM(2, 5) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 6))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 6) + (TaxableAmtSta# * (StateTax(1).STM(2, 6) * 0.01))) / PayFreq)
'      End If
'
'
'    Case "H", ""  'Georgia table H,"S", - Single Individual
'
'      '--Get Standard Deduction
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX301)
'
'      '--Get personal and dependendant deductions if claiming at least one allowance
'      If Emp2Rec(1).EMPSTAA > 0 Then
'        '--Personal Allowance Amount
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX302)
'        '--Dependent Exemption
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX303))
'      End If
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# <= StateTax(1).STH(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 1) And TaxableAmtSta# <= StateTax(1).STH(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 2) And TaxableAmtSta# <= StateTax(1).STH(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 3) And TaxableAmtSta# <= StateTax(1).STH(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 3) + (TaxableAmtSta# * (StateTax(1).STH(2, 3) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 4) And TaxableAmtSta# <= StateTax(1).STH(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 4) + (TaxableAmtSta# * (StateTax(1).STH(2, 4) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 5) And TaxableAmtSta# <= StateTax(1).STH(3, 6) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 5) + (TaxableAmtSta# * (StateTax(1).STH(2, 5) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 6))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 6) + (TaxableAmtSta# * (StateTax(1).STH(2, 6) * 0.01))) / PayFreq)
'      End If
'
'  End Select
'
'  '--Add additional set withholding amount
'  If Emp2Rec(1).EMPSTAAA > 0 Then
'    TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'  End If
'
'
'  Case "AR"
'  'add some kind of an include section here for various state calc's
''look at employer's state in the controll file to determin which
''state tax tables to use or edit
'
''********Early run throughs got an overflow error/divide by zero
''********error in the "M" section with TAX205. Need an "On Error"
''********statement in this sub when debugging for final errors
'
''  If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
''
''  Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
''
''    Case "S", ""
''
''      AR10PctDed# = OldRound#(AnnualizedStaGross# * StateTax(1).TAX105)
''
''      If AR10PctDed# > StateTax(1).TAX103 Then
''        AR10PctDed# = StateTax(1).TAX103
''      End If
''
''      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
''
''      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
''
''      If TaxableAmtSta# <= 0 Then
''        TransRec.StaTaxAmt = 0
''      ElseIf TaxableAmtSta# < StateTax(1).STS(3, 1) Then
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# < StateTax(1).STS(3, 2) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 3) And TaxableAmtSta# < StateTax(1).STS(3, 4) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 3))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 4) And TaxableAmtSta# < StateTax(1).STS(3, 5) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 4))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * 0.01))))
''
''      Else
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 5))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * 0.01))))
''      End If
''
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX106)
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX107))
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
''
''   Case "M"
''
''      AR10PctDed# = OldRound#(AnnualizedStaGross# / StateTax(1).TAX205)
''
''      If AR10PctDed# > StateTax(1).TAX203 Then
''        AR10PctDed# = StateTax(1).TAX203
''      End If
''
''      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
''      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
''
''      If TaxableAmtSta# <= 0 Then
''        TransRec.StaTaxAmt = 0
''
''      ElseIf TaxableAmtSta# < StateTax(1).STM(3, 1) Then
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 1) And TaxableAmtSta# < StateTax(1).STM(3, 2) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 1))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 2) And TaxableAmtSta# < StateTax(1).STM(3, 3) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 2))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 3) And TaxableAmtSta# < StateTax(1).STM(3, 4) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 3))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 4) And TaxableAmtSta# < StateTax(1).STM(3, 5) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 4))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 5) + (TaxableAmtSta# * (StateTax(1).STM(2, 5) * 0.01))))
''
''      Else
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 5))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 6) + (TaxableAmtSta# * (StateTax(1).STM(2, 6) * 0.01))))
''      End If
''
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX206)
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX207))
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
''
''    Case "H"
''
''      AR10PctDed# = OldRound(AnnualizedStaGross# * StateTax(1).TAX305)
''
''      If AR10PctDed# > StateTax(1).TAX303 Then
''        AR10PctDed# = StateTax(1).TAX303
''      End If
''
''      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
''      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
''
''      If TaxableAmtSta# <= 0 Then
''        TransRec.StaTaxAmt = 0
''
''      ElseIf TaxableAmtSta# < StateTax(1).STH(3, 1) Then
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 1) And TaxableAmtSta# < StateTax(1).STH(3, 2) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 1))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 2))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 3) + (TaxableAmtSta# * (StateTax(1).STH(2, 3) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 3) And TaxableAmtSta# < StateTax(1).STH(3, 4) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 3))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 4) + (TaxableAmtSta# * (StateTax(1).STH(2, 4) * 0.01))))
''
''      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 4) And TaxableAmtSta# < StateTax(1).STH(3, 5) Then
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 4))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 5) + (TaxableAmtSta# * (StateTax(1).STH(2, 5) * 0.01))))
''
''      Else
''        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 5))
''        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 6) + (TaxableAmtSta# * (StateTax(1).STH(2, 6) * 0.01))))
''      End If
''
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX306)
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX307))
''      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
''
''  End Select
''
''  'deduct tax credit from state tax amt.
''  If TransRec.StaTaxAmt < 0 Then
''    TransRec.StaTaxAmt = 0
''  End If
''
''  TaxableAmtSta# = TransRec.StaTaxAmt
''
'''  End Select
''add some kind of an include section here for various state calc's
''look at employer's state in the controll file to determin which
''state tax tables to use or edit
'
''Arkansas 1/21/97
'
'  If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'  Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
'
'    Case "S", ""
'
'      'AR10PctDed# = OldRound(AnnualizedStaGross# * StateTax(1).TAX105)
'
'      'IF AR10PctDed# > StateTax(1).TAX103 THEN
'        AR10PctDed# = StateTax(1).TAX103
'      'END IF
'
'      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# < StateTax(1).STS(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# < StateTax(1).STS(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 3) And TaxableAmtSta# < StateTax(1).STS(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 4) And TaxableAmtSta# < StateTax(1).STS(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * 0.01))))
'
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * 0.01))))
'      End If
''STOP
'      'TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX106)
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX106))
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
'
'   Case "M"
'
'      'AR10PctDed# = OldRound(AnnualizedStaGross# / StateTax(1).TAX205)
'      'IF AR10PctDed# > StateTax(1).TAX203 THEN
'        AR10PctDed# = StateTax(1).TAX203
'      'END IF
'
'      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'
'      ElseIf TaxableAmtSta# < StateTax(1).STM(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 1) And TaxableAmtSta# < StateTax(1).STM(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 2) And TaxableAmtSta# < StateTax(1).STM(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 3) And TaxableAmtSta# < StateTax(1).STM(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 4) And TaxableAmtSta# < StateTax(1).STM(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 5) + (TaxableAmtSta# * (StateTax(1).STM(2, 5) * 0.01))))
'
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 6) + (TaxableAmtSta# * (StateTax(1).STM(2, 6) * 0.01))))
'      End If
'
'      'TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX206)
'      'TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX207))
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX206))
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
'
'    Case "H"
'
'      'AR10PctDed# = OldRound(AnnualizedStaGross# * StateTax(1).TAX305)
'      'IF AR10PctDed# > StateTax(1).TAX303 THEN
'        AR10PctDed# = StateTax(1).TAX303
'      'END IF
'
'      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross#)
'
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'
'      ElseIf TaxableAmtSta# < StateTax(1).STH(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 1) And TaxableAmtSta# < StateTax(1).STH(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 3) + (TaxableAmtSta# * (StateTax(1).STH(2, 3) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 3) And TaxableAmtSta# < StateTax(1).STH(3, 4) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 4) + (TaxableAmtSta# * (StateTax(1).STH(2, 4) * 0.01))))
'
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 4) And TaxableAmtSta# < StateTax(1).STH(3, 5) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 4))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 5) + (TaxableAmtSta# * (StateTax(1).STH(2, 5) * 0.01))))
'
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 5))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 6) + (TaxableAmtSta# * (StateTax(1).STH(2, 6) * 0.01))))
'      End If
'
'      'TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - StateTax(1).TAX306)
'      'TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX307))
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX306))
'      TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt / PayFreq)
'
'  End Select
'
'  'deduct tax credit from state tax amt.
'  If TransRec.StaTaxAmt < 0 Then
'    TransRec.StaTaxAmt = 0
'  End If
'
'
'  If Emp2Rec(1).EMPSTAAA > 0 Then
'    TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'  End If
'
'  TaxableAmtSta# = TransRec.StaTaxAmt
'
''
''  END SELECT
'
'
'
'  Case "AL"
'
'  'add some kind of an include section here for various state calc's
'  'look at employer's state in the controll file to determin which
'  'state tax tables to use or edit
'
'  'Alabama 1/8/97
'
'  AL20PctDed# = OldRound(AnnualizedStaGross# * 0.2)
'  FedTaxDed# = TransRec.FedTaxAmt * PayFreq
'
'  'STOP
'
'  Select Case QPTrim$(Emp2Rec(1).EMPSTAS)
'
'    Case "S", ""
'      If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'      '--new for al state calc
'      If AL20PctDed# >= 2000 Then AL20PctDed# = 2000
'      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
'      '--end of al specific
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX106))
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX103)
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# < StateTax(1).STS(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 1) And TaxableAmtSta# < StateTax(1).STS(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq)
'      End If
'
'    Case "M"
'      If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'      '--new for al state calc
'      If AL20PctDed# >= 4000 Then AL20PctDed# = 4000
'      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
'      '--end of al specific
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX206))
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX203)
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# < StateTax(1).STM(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 1) And TaxableAmtSta# < StateTax(1).STM(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STM(3, 2) And TaxableAmtSta# < StateTax(1).STM(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STM(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * 0.01))) / PayFreq)
'      End If
'
'    Case "H"
'      If Emp2Rec(1).EMPSTAA < 0 Then Emp2Rec(1).EMPSTAA = 0
'
'      '--new for al state calc
'      If AL20PctDed# >= 2000 Then AL20PctDed# = 2000
'      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
'      '--end of al specific
'
'      TaxableAmtSta# = OldRound#(AnnualizedStaGross# - (Emp2Rec(1).EMPSTAA * StateTax(1).TAX306))
'      TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).TAX303)
'      If TaxableAmtSta# <= 0 Then
'        TransRec.StaTaxAmt = 0
'      ElseIf TaxableAmtSta# < StateTax(1).STH(3, 1) Then
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STH(3, 1) And TaxableAmtSta# < StateTax(1).STH(3, 2) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STH(3, 1))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * 0.01))) / PayFreq)
'      ElseIf TaxableAmtSta# > StateTax(1).STS(3, 2) And TaxableAmtSta# < StateTax(1).STS(3, 3) Then
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 2))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * 0.01))) / PayFreq)
'      Else
'        TaxableAmtSta# = OldRound#(TaxableAmtSta# - StateTax(1).STS(3, 3))
'        TransRec.StaTaxAmt = OldRound#(OldRound#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * 0.01))) / PayFreq)
'      End If
'  End Select
'
'  If Emp2Rec(1).EMPSTAAA > 0 Then
'    TransRec.StaTaxAmt = OldRound#(TransRec.StaTaxAmt + Emp2Rec(1).EMPSTAAA)
'  End If
'
'
'  Case Else
'
'  End Select
'  Return
'End Sub
'
'Public Sub TaxTextLoad()
'   Dim UnitHandle As Integer
'   Dim UnitFileRec As UnitFileRecType
'   OpenUnitFile UnitHandle
'   Get UnitHandle, 1, UnitFileRec
'   Close UnitHandle
'   Select Case UnitFileRec.UFSTATE
'   Case "GA":
'     TaxText$(1) = "GA"
'     TaxText$(2) = "GASingle"
'     TaxText$(3) = "GAMarried"
'     TaxText$(4) = "GAHead of Household"
'   Case "SC":
'     TaxText$(1) = "SC"
'   Case "OK":
'     TaxText$(1) = "OK"
'     TaxText$(2) = "Single"
'     TaxText$(3) = "Married, Head of Household"
'     TaxText$(4) = "Dual Income Married"
'   Case "AR":
'     TaxText$(1) = "AR"
'     TaxText$(2) = "Single"
'     TaxText$(3) = "Married (1 Exempt'n)"
'     TaxText$(4) = "Married/Head Fam(2 Exempt'ns"
'   Case "AL":
'     TaxText$(1) = "AL"
'     TaxText$(2) = "Single"
'     TaxText$(3) = "Married"
'     TaxText$(4) = "Head of Household"
'   Case "VA":
'     TaxText$(1) = "VA"
'     TaxText$(2) = "Single"
'     TaxText$(3) = "Married"
'     TaxText$(4) = "Head of Household"
'   Case Else:
'     TaxText$(1) = "NC"
'     TaxText$(2) = "Single"
'     TaxText$(3) = "Married"
'     TaxText$(4) = "Head of Household"
'   End Select
'
'End Sub
'
'Public Sub ParseTrans2Hourly(TransRec() As TransRecType, HourInput() As HourlyInputType)
'  Dim Cnt As Integer
'  For Cnt = 1 To 8
'    HourInput(1).HDist(Cnt).DAcct = TransRec(1).TDist(Cnt).DAcct
'    HourInput(1).HDist(Cnt).DRHrs = TransRec(1).TDist(Cnt).DRHrs
'    HourInput(1).HDist(Cnt).DOHrs = TransRec(1).TDist(Cnt).DOHrs
'  Next
'
'  HourInput(1).WORKHRS = TransRec(1).RegHrsWork
'  HourInput(1).VACHRS = TransRec(1).VacUsed
'  HourInput(1).SICKHRS = TransRec(1).SickUsed
'  HourInput(1).HOLHRS = TransRec(1).HOLHOURS
'  HourInput(1).COMPHRS = TransRec(1).CompUsed
'
'  HourInput(1).PerHRS = TransRec(1).PerHours
'
'  HourInput(1).TOTHRSPD = TransRec(1).RegHrsPaid
'  HourInput(1).OTWORKED = TransRec(1).OTHours
'  HourInput(1).OTHRSPD = TransRec(1).OTHrsPaid
'  HourInput(1).OT2Comp = TransRec(1).OT2Comp
'
'  HourInput(1).ALTEARN1 = TransRec(1).EAmt(1)   '*
'  HourInput(1).ALTEARN2 = TransRec(1).EAmt(2)   'actual earning amounts
'  HourInput(1).ALTEARN3 = TransRec(1).EAmt(3)   '*
'
'  HourInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct            '*
'  HourInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct            '
'  HourInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct            'Earnings distribution accts.
'
'  HourInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct            '
'  HourInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct            '
'  HourInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct            '*
'
'  HourInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt             '*
'  HourInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt             '
'  HourInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt             'earnings amounts as distributed
'  HourInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt             'to accts.
'  HourInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt             '
'  HourInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt             '*
'
'  HourInput(1).TOTAERN = TransRec(1).TotAdditEarn
'  HourInput(1).TaxFring = TransRec(1).TaxFring
'
'End Sub
'
'Sub ParseTrans2Salary(TransRec() As TransRecType, SalInput() As SalaryInputType)
'  Dim x As Integer
'
'  If Len(QPTrim$(TransRec(1).PaySFlag)) = 0 Then
'    SalInput(1).PAYSAL = "Y"
'  Else
'    SalInput(1).PAYSAL = TransRec(1).PaySFlag
'  End If
'
'  SalInput(1).VACHRS = TransRec(1).VacUsed
'  SalInput(1).SICKHRS = TransRec(1).SickUsed
'  SalInput(1).HOLHRS = TransRec(1).HOLHOURS
'  SalInput(1).PerHRS = TransRec(1).PerHours
'  SalInput(1).COMPHRS = TransRec(1).CompUsed 'added 9/1/04
'  SalInput(1).Hrs2Cmp = TransRec(1).OT2Comp 'added 9/1/04
'
'  For x = 1 To 8
'    SalInput(1).SDist(x).DAcct = TransRec(1).TDist(x).DAcct
'    SalInput(1).SDist(x).DPct = TransRec(1).TDist(x).DPct
'  Next
'
'  SalInput(1).ALTEARN1 = TransRec(1).EAmt(1)
'  SalInput(1).ALTEARN2 = TransRec(1).EAmt(2)
'  SalInput(1).ALTEARN3 = TransRec(1).EAmt(3)
'
'  SalInput(1).AERNAMT1 = TransRec(1).EDist(1).EAmt
'  SalInput(1).AERNAMT2 = TransRec(1).EDist(2).EAmt
'  SalInput(1).AERNAMT3 = TransRec(1).EDist(3).EAmt
'  SalInput(1).AERNAMT4 = TransRec(1).EDist(4).EAmt
'  SalInput(1).AERNAMT5 = TransRec(1).EDist(5).EAmt
'  SalInput(1).AERNAMT6 = TransRec(1).EDist(6).EAmt
'
'  SalInput(1).AERNDST1 = TransRec(1).EDist(1).EAcct
'  SalInput(1).AERNDST2 = TransRec(1).EDist(2).EAcct
'  SalInput(1).AERNDST3 = TransRec(1).EDist(3).EAcct
'  SalInput(1).AERNDST4 = TransRec(1).EDist(4).EAcct
'  SalInput(1).AERNDST5 = TransRec(1).EDist(5).EAcct
'  SalInput(1).AERNDST6 = TransRec(1).EDist(6).EAcct
'  SalInput(1).TOTAERN = TransRec(1).TotAdditEarn
'  SalInput(1).TaxFring = TransRec(1).TaxFring
'
'End Sub
'
'Sub ParseHourly2Trans(TransRec() As TransRecType, HourInput() As HourlyInputType)
'  Dim Cnt As Integer
'
'  For Cnt = 1 To 8
'    TransRec(1).TDist(Cnt).DAcct = HourInput(1).HDist(Cnt).DAcct
'    TransRec(1).TDist(Cnt).DRHrs = HourInput(1).HDist(Cnt).DRHrs
'    TransRec(1).TDist(Cnt).DOHrs = HourInput(1).HDist(Cnt).DOHrs
'  Next
'
'  TransRec(1).RegHrsWork = HourInput(1).WORKHRS
'  TransRec(1).VacUsed = HourInput(1).VACHRS
'  TransRec(1).SickUsed = HourInput(1).SICKHRS
'  TransRec(1).HOLHOURS = HourInput(1).HOLHRS
'  TransRec(1).CompUsed = HourInput(1).COMPHRS
'
'  TransRec(1).PerHours = HourInput(1).PerHRS
'
'  TransRec(1).RegHrsPaid = HourInput(1).TOTHRSPD
'  TransRec(1).OTHours = HourInput(1).OTWORKED
'  TransRec(1).OTHrsPaid = HourInput(1).OTHRSPD
'  TransRec(1).OT2Comp = HourInput(1).OT2Comp
'
'  TransRec(1).EAmt(1) = HourInput(1).ALTEARN1
'  TransRec(1).EAmt(2) = HourInput(1).ALTEARN2
'  TransRec(1).EAmt(3) = HourInput(1).ALTEARN3
'  TransRec(1).TotAdditEarn = HourInput(1).TOTAERN 'this is sum of (alt earn) fields
'
'  TransRec(1).EDist(1).EAcct = HourInput(1).AERNDST1
'  TransRec(1).EDist(2).EAcct = HourInput(1).AERNDST2
'  TransRec(1).EDist(3).EAcct = HourInput(1).AERNDST3
'  TransRec(1).EDist(4).EAcct = HourInput(1).AERNDST4
'  TransRec(1).EDist(5).EAcct = HourInput(1).AERNDST5
'  TransRec(1).EDist(6).EAcct = HourInput(1).AERNDST6
'
'  TransRec(1).EDist(1).EAmt = HourInput(1).AERNAMT1
'  TransRec(1).EDist(2).EAmt = HourInput(1).AERNAMT2
'  TransRec(1).EDist(3).EAmt = HourInput(1).AERNAMT3
'  TransRec(1).EDist(4).EAmt = HourInput(1).AERNAMT4
'  TransRec(1).EDist(5).EAmt = HourInput(1).AERNAMT5
'  TransRec(1).EDist(6).EAmt = HourInput(1).AERNAMT6
'
'  TransRec(1).TaxFring = HourInput(1).TaxFring
'End Sub
'
'Sub ParseSalary2Trans(TransRec() As TransRecType, SalInput() As SalaryInputType)
'  Dim Cnt As Integer
'
'  TransRec(1).PaySFlag = SalInput(1).PAYSAL
'
'  TransRec(1).VacUsed = SalInput(1).VACHRS
'  TransRec(1).SickUsed = SalInput(1).SICKHRS
'  TransRec(1).HOLHOURS = SalInput(1).HOLHRS
'  TransRec(1).PerHours = SalInput(1).PerHRS
'  TransRec(1).CompUsed = SalInput(1).COMPHRS 'added 9/1/04
'  TransRec(1).OT2Comp = SalInput(1).Hrs2Cmp 'added 9/1/04
'
'  For Cnt = 1 To 8
'    TransRec(1).TDist(Cnt).DAcct = SalInput(1).SDist(Cnt).DAcct
'    TransRec(1).TDist(Cnt).DPct = SalInput(1).SDist(Cnt).DPct
'  Next
'
'  TransRec(1).EAmt(1) = SalInput(1).ALTEARN1
'  TransRec(1).EAmt(2) = SalInput(1).ALTEARN2
'  TransRec(1).EAmt(3) = SalInput(1).ALTEARN3
'
'  TransRec(1).EDist(1).EAmt = SalInput(1).AERNAMT1
'  TransRec(1).EDist(2).EAmt = SalInput(1).AERNAMT2
'  TransRec(1).EDist(3).EAmt = SalInput(1).AERNAMT3
'  TransRec(1).EDist(4).EAmt = SalInput(1).AERNAMT4
'  TransRec(1).EDist(5).EAmt = SalInput(1).AERNAMT5
'  TransRec(1).EDist(6).EAmt = SalInput(1).AERNAMT6
'
'  TransRec(1).EDist(1).EAcct = SalInput(1).AERNDST1
'  TransRec(1).EDist(2).EAcct = SalInput(1).AERNDST2
'  TransRec(1).EDist(3).EAcct = SalInput(1).AERNDST3
'  TransRec(1).EDist(4).EAcct = SalInput(1).AERNDST4
'  TransRec(1).EDist(5).EAcct = SalInput(1).AERNDST5
'  TransRec(1).EDist(6).EAcct = SalInput(1).AERNDST6
'
'  TransRec(1).TotAdditEarn = SalInput(1).TOTAERN
'  TransRec(1).TaxFring = SalInput(1).TaxFring
'
'End Sub
'
'Sub ParseTrans2ScrnCalc(TransRec() As TransRecType, ScrnCalc() As ScrnCalcType)
'  'STOP
'  Dim x As Integer
'
''  TransRec(1).TDist(1).DRWage = TransRec(1).TDist(1).DRWage
'
'  ScrnCalc(1).REGEARN = TransRec(1).TotRegWage
'  ScrnCalc(1).OTEARN = TransRec(1).TotOTWage
'
'  ScrnCalc(1).ALTEARN1 = TransRec(1).EAmt(1)
'  ScrnCalc(1).ALTEARN2 = TransRec(1).EAmt(2)
'  ScrnCalc(1).ALTEARN3 = TransRec(1).EAmt(3)
'
'  ScrnCalc(1).SOCTAX = TransRec(1).SocTaxAmt
'  ScrnCalc(1).MEDTAX = TransRec(1).MedTaxAmt
'  ScrnCalc(1).FEDTAX = TransRec(1).FedTaxAmt
'  ScrnCalc(1).STATAX = TransRec(1).StaTaxAmt
'  ScrnCalc(1).RETIRE = TransRec(1).RetireAmt
'
'  For x = 1 To 50
'   ScrnCalc(1).Ded(x) = TransRec(1).DAmt(x)
'  Next x
'
'  ScrnCalc(1).GrossPay = TransRec(1).GrossPay
'
'  ScrnCalc(1).TOTDED = OldRound#(TransRec(1).TotDedAmt + TransRec(1).TotTaxAmt)
'
'  ScrnCalc(1).EIC = TransRec(1).EICAmt
'  ScrnCalc(1).NetPay = TransRec(1).NetPay
'End Sub
'
'Sub ParseScrnCalc2Trans(TransRec() As TransRecType, ScrnCalc() As ScrnCalcType)
'
'  Dim Cnt As Integer
'
'  TransRec(1).TotRegWage = ScrnCalc(1).REGEARN
'  TransRec(1).TotOTWage = ScrnCalc(1).OTEARN
'
'  TransRec(1).EAmt(1) = ScrnCalc(1).ALTEARN1
'  TransRec(1).EAmt(2) = ScrnCalc(1).ALTEARN2
'  TransRec(1).EAmt(3) = ScrnCalc(1).ALTEARN3
'
'  TransRec(1).SocTaxAmt = ScrnCalc(1).SOCTAX
'  TransRec(1).MedTaxAmt = ScrnCalc(1).MEDTAX
'  TransRec(1).FedTaxAmt = ScrnCalc(1).FEDTAX
'  TransRec(1).StaTaxAmt = ScrnCalc(1).STATAX
'
'  TransRec(1).RetireAmt = ScrnCalc(1).RETIRE
'
'  For Cnt = 1 To 50
'    TransRec(1).DAmt(Cnt) = ScrnCalc(1).Ded(Cnt)
'  Next Cnt
'
'  TransRec(1).GrossPay = ScrnCalc(1).GrossPay
'
'  TransRec(1).TotDedAmt = 0
'  For Cnt = 1 To 50
'    TransRec(1).TotDedAmt = OldRound#(TransRec(1).TotDedAmt + TransRec(1).DAmt(Cnt))
'  Next
'  'fix from region-d
'  TransRec(1).TotDedAmt = OldRound#(TransRec(1).TotDedAmt + TransRec(1).RetireAmt)
'
'  TransRec(1).EICAmt = ScrnCalc(1).EIC
'  TransRec(1).NetPay = ScrnCalc(1).NetPay
'
'
'End Sub
'
'Function CheckFor2ManyDecimals(Text As String) As Boolean
'  Dim Cnt As Integer
'  Dim DecCnt As Integer
'  Dim StrLen As Long
'  Dim thischar$
'  'this function traps errors created when a user keys in a
'  'decimal value and inadvertantly keys in more than 1 decimal
'  StrLen = Len(Text)
'  For Cnt = 1 To StrLen
'    thischar = Mid$(Text, Cnt, 1)
'    If thischar = "." Then DecCnt = DecCnt + 1 'counts decimals
'  Next Cnt
'  If DecCnt > 1 Then 'if decimal count is more than 1 then process
'  'accordingly
'    CheckFor2ManyDecimals = True
'  Else
'    CheckFor2ManyDecimals = False
'  End If
'End Function
'
''Public Sub PostTransactions()
''
''  Dim FirstFlag As Boolean
''  Dim cnt&, cnt1&
''  Dim NumOfRecs&
''  Dim CHandle As Integer
''  Dim zz As Integer
''  Dim OHandle As Integer, GLDate As Long
''  Dim OSNumRec As Integer, GLIFDate$
''  Dim PostDate As Long
''  Dim TransRecLen As Integer
''  Dim Emp2RecLen As Integer
''  Dim Emp3RecLen As Integer
''  Dim CheckRecLen As Integer
''  Dim OSChkRecLen As Integer
''  Dim PSysHandle As Integer
''  Dim EHandle2 As Integer
''  Dim EHandle3 As Integer
''  Dim THandle As Integer
''  Dim HHandle As Integer
''  Dim GLSetUpRecLen As Integer, GLSetUpName$
''  Dim PDRLen As Integer, OSChkFile$
''  Dim DoOSChkFlag As Boolean
''  Dim GLHandle As Integer
''  Dim NextHistRec&
''  Dim GLIFRecLen As Integer
''  Dim GLRecLen As Integer, NextRec&
''  Dim IHandle As Integer, PRIF$
''  Dim PPDHandle As Integer
''  Dim BadAccts As Integer
''  '----added 6/17/04-----------
''  Dim TempVoid As VoidCheckType
''  Dim TVHandle As Integer
''  Dim NumOfTempVoids As Integer
''  Dim VCnt As Integer
''  '----added 6/17/04---^^^-----
''
''  FirstFlag = True
''
''  PostDate = Date2Num(Date$)
''
''  ReDim TransRec(1) As TransRecType
''  ReDim EmpRec2(1) As EmpData2Type
''  ReDim EmpRec3(1) As EmpData3Type
''  ReDim Check(1) As PRCheckRecType
''
''  ReDim PSysRec(1) As RegDSysFileRecType
''
''  ReDim OSIFRec(1) As OSChekRecType
''  ReDim PrdDefRecHere(1) As PeriodDefaultRecType
''  ReDim PDR(1) As PeriodDefaultRecType
''
''  TransRecLen = Len(TransRec(1))
''  Emp2RecLen = Len(EmpRec2(1))
''  Emp3RecLen = Len(EmpRec3(1))
''  CheckRecLen = Len(Check(1))
''  OSChkRecLen = Len(OSIFRec(1))
''
''  OpenSysFile PSysHandle
''  Get PSysHandle, 1, PSysRec(1)
''
''  Close PSysHandle
''  ReDim GLSetupRec(1) As GLSetupRecType
''
''  GLSetUpRecLen = Len(GLSetupRec(1))
''
'''  GLSetUpName$ = QPTrim$(PSysRec(1).CITIDIR) + "\GLSETUP.DAT"
''  If Mid(CurrCitiPath, Len(CurrCitiPath), 1) <> "\" Then
''    GLSetUpName$ = QPTrim$(CurrCitiPath + "\GLSETUP.DAT")
''  ElseIf Mid(CurrCitiPath, Len(CurrCitiPath), 1) = "\" Then
''    GLSetUpName$ = QPTrim$(CurrCitiPath + "GLSETUP.DAT")
''  End If
''
''  GLHandle = FreeFile
''
''  If Exist(GLSetUpName$) Then
''    Open GLSetUpName$ For Random Shared As GLHandle Len = GLSetUpRecLen
''    Get GLHandle, 1, GLSetupRec(1)
''  End If
''  Close GLHandle
''
''
''  '05/10/94   Fix to set false after posting
''  '    -vvv-  is used as a flag to menu entry points
''  '06/09/94   Fix for manual trans
''
''  'Note: the program cannot get this far if both MACTIVE
''  ' and PACTIVE are false...it is trapped out
''  PDRLen = Len(PDR(1))
''  OpenPPDefaultFile PPDHandle
''  Get PPDHandle, 1, PDR(1)
''  If PDR(1).MACTIVE = True Then
''    EntryType = 2
''  ElseIf PDR(1).PACTIVE = True Then
''    EntryType = 1
''  End If
''
''  Select Case EntryType
''  Case 1 'Normal
''    PrdDefRecHere(1).PACTIVE = False
''    PDR(1).PACTIVE = False      'set active flag = false
''  Case 2 'Manual
''    PrdDefRecHere(1).MACTIVE = False
''    PDR(1).MACTIVE = False      'set active flag = false
''  End Select
''
''  Put PPDHandle, 1, PDR(1)
''  Close PPDHandle
''
'' 'open checks data file
''  OpenChecksFile CHandle
''  NumOfRecs = LOF(CHandle) / Len(Check(1))
''  OpenEmpData2File EHandle2
''  OpenEmpData3File EHandle3
''  OpenTransWorkFile THandle
''
''  OpenTransHistFile HHandle
''  NextHistRec& = LOF(HHandle) / Len(TransRec(1)) + 1
''
''  OpenTempVoidFile TVHandle '6/17/04
''  NumOfTempVoids = LOF(TVHandle) / Len(TempVoid)
''
'''  OSChkFile$ = QPTrim$(PSysRec(1).CITIDIR)
''  OSChkFile$ = CurrCitiPath
''  If Len(OSChkFile$) > 0 Then
''    DoOSChkFlag = True
''  End If
''  If DoOSChkFlag Then
''    If Right$(OSChkFile$, 1) <> "\" Then
''      OSChkFile$ = OSChkFile$ + "\CRCHK.DAT"
''    Else
''      OSChkFile$ = OSChkFile$ + "CRCHK.DAT"
''    End If
''    Call OpenOSChekFile(OHandle, OSNumRec)
''    If OSNumRec = 0 Then
''      OSNumRec = 1
''    Else 'added the "else" on 10/13/03
''      OSNumRec = OSNumRec + 1
''    End If
''  End If
''
''For cnt& = 1 To NumOfRecs&
''  Get CHandle, cnt&, Check(1)
''  Get THandle, cnt, TransRec(1)
''
''  If TransRec(1).TActive = -1 Then
''    'if this is an active check or transaction
''    If FirstFlag Then
''      FirstFlag = False
''      GLIFDate$ = MakeRegDate(Check(1).CheckDate)
''        GLDate = Check(1).CheckDate
''        ReplaceString GLIFDate$, "/", "-"
''        For zz = 1 To 9
''          ReplaceString GLIFDate$, "200" + QPTrim$(Str$(zz)), "0" + QPTrim$(Str$(zz))
''        Next
''
''        GLIFDate$ = QPTrim$(GLIFDate$)
''      End If
''      '** Get all data from files that need to be updated
''      Get EHandle2, cnt, EmpRec2(1)
''      Get EHandle3, cnt, EmpRec3(1)
''      '** Update employee 2 file and adjust previous transaction pointer
''      If EmpRec2(1).EMPVACE < 0 Then EmpRec2(1).EMPVACE = 0
''      If EmpRec2(1).EMPSLE < 0 Then EmpRec2(1).EMPSLE = 0
''      If EmpRec2(1).EMPCTE < 0 Then EmpRec2(1).EMPCTE = 0
''
''      If EmpRec2(1).EMPVUSED < 0 Then EmpRec2(1).EMPVUSED = 0
''      If EmpRec2(1).EMPSLUSE < 0 Then EmpRec2(1).EMPSLUSE = 0
''      If EmpRec2(1).EMPCTUSE < 0 Then EmpRec2(1).EMPCTUSE = 0
''
''      EmpRec2(1).EMPVUSED = OldRound(EmpRec2(1).EMPVUSED + TransRec(1).VacUsed)
''      EmpRec2(1).EMPVBAL = OldRound(EmpRec2(1).EMPVACE - EmpRec2(1).EMPVUSED)
''
''      EmpRec2(1).EMPSLUSE = OldRound(EmpRec2(1).EMPSLUSE + TransRec(1).SickUsed)
''      EmpRec2(1).EMPSLBAL = OldRound(EmpRec2(1).EMPSLE - EmpRec2(1).EMPSLUSE)
''
''      '022498
''      EmpRec2(1).HolUsed = OldRound(EmpRec2(1).HolUsed + TransRec(1).HOLHOURS)
''      EmpRec2(1).HOLBAL = OldRound(EmpRec2(1).HOLERN - EmpRec2(1).HolUsed)
''
''      EmpRec2(1).PerUsed = OldRound(EmpRec2(1).PerUsed + TransRec(1).PerHours)
''      EmpRec2(1).PERBAL = OldRound(EmpRec2(1).PERERN - EmpRec2(1).PerUsed)
''      '***************
''      'fix for comp earned 05/04/94
''      EmpRec2(1).EMPCTUSE = OldRound(EmpRec2(1).EMPCTUSE + TransRec(1).CompUsed)
''
''      If EntryType = Normal Then
''        EmpRec2(1).EMPCTBAL = Check(1).CompBal
''        EmpRec2(1).EMPCTE = Check(1).CompEarn
''      Else
''        EmpRec2(1).EMPCTBAL = OldRound(EmpRec2(1).EMPCTE - EmpRec2(1).EMPCTUSE)
''      End If
''
''
''      'adjust and update (last - previous) transaction pointers
''      If EmpRec2(1).LastTransRec >= 0 Then
''        TransRec(1).PrevTransRec = EmpRec2(1).LastTransRec
''      Else
''        TransRec(1).PrevTransRec = 0
''      End If
''      EmpRec2(1).LastTransRec = CInt(NextHistRec&)
''
''      '** Update employee 3 file
''      'EmpRec3(1).YTDGrossPay = OldRound(EmpRec3(1).YTDGrossPay + Check(1).GrossPay)
''      '-=-=man
''      EmpRec3(1).YTDGrossPay = OldRound(EmpRec3(1).YTDGrossPay + TransRec(1).GrossPay)
''      EmpRec3(1).YTDFedGrossPay = OldRound(EmpRec3(1).YTDFedGrossPay + TransRec(1).FedGrossPay)
''      EmpRec3(1).YTDStaGrossPay = OldRound(EmpRec3(1).YTDStaGrossPay + TransRec(1).StaGrossPay)
''      EmpRec3(1).YTDSocGrossPay = OldRound(EmpRec3(1).YTDSocGrossPay + TransRec(1).SocGrossPay)
''      EmpRec3(1).YTDMedGrossPay = OldRound(EmpRec3(1).YTDMedGrossPay + TransRec(1).MedGrossPay)
''
''      EmpRec3(1).YTDRegPay = OldRound(EmpRec3(1).YTDRegPay + TransRec(1).TotRegWage)
''      EmpRec3(1).YTDOTPay = OldRound(EmpRec3(1).YTDOTPay + TransRec(1).TotOTWage)
''
''      EmpRec3(1).YTDNet = OldRound(EmpRec3(1).YTDNet + TransRec(1).NetPay)
''
''      EmpRec3(1).YTDFederal = OldRound(EmpRec3(1).YTDFederal + TransRec(1).FedTaxAmt)
''      EmpRec3(1).YTDState = OldRound(EmpRec3(1).YTDState + TransRec(1).StaTaxAmt)
''      EmpRec3(1).YTDSocial = OldRound(EmpRec3(1).YTDSocial + TransRec(1).SocTaxAmt)
''      EmpRec3(1).YTDMedicare = OldRound(EmpRec3(1).YTDMedicare + TransRec(1).MedTaxAmt)
''      EmpRec3(1).YTDRetire = OldRound(EmpRec3(1).YTDRetire + TransRec(1).RetireAmt)
''
''      'year to date totals on deductions
''      For cnt1 = 1 To 50
''        EmpRec3(1).YTDDAmt(cnt1) = OldRound(EmpRec3(1).YTDDAmt(cnt1) + TransRec(1).DAmt(cnt1))
''        EmpRec3(1).YTDDAmtT = OldRound(EmpRec3(1).YTDDAmtT + TransRec(1).DAmt(cnt1))
''      Next
''
''      'year to date totals on alt earnings
''      EmpRec3(1).YTDEarn1 = OldRound(EmpRec3(1).YTDEarn1 + TransRec(1).EAmt(1))
''      EmpRec3(1).YTDEarn2 = OldRound(EmpRec3(1).YTDEarn2 + TransRec(1).EAmt(2))
''      EmpRec3(1).YTDEarn3 = OldRound(EmpRec3(1).YTDEarn3 + TransRec(1).EAmt(3))
''      EmpRec3(1).YTDEarnT = OldRound(EmpRec3(1).YTDEarn1 + EmpRec3(1).YTDEarn2 + EmpRec3(1).YTDEarn3)
''
''      '** Update MISC transaction history data
''      'added fix for manual transaction entry 5/13/94  friday the 13th
''      If EntryType = Normal Then
''        TransRec(1).CheckNum = Check(1).CheckNum
''        TransRec(1).CheckDate = Check(1).CheckDate
''        TransRec(1).PostDate = PostDate
''      End If
''
''      '** Added Update EIC year to date.  6/06/94
''      If TransRec(1).EICAmt > 0 Then
''        EmpRec3(1).YTDEIC = OldRound(EmpRec3(1).YTDEIC + TransRec(1).EICAmt)
''      End If
''
''      '** Update active flags
''      TransRec(1).TActive = False
''      Check(1).CActive = False
''      '** Update OSChk File
''      If DoOSChkFlag Then
''        OSIFRec(1).Chknum = Check(1).CheckNum
''        OSIFRec(1).ChkDate = Date2Num(GLIFDate$)
''        OSIFRec(1).Desc = Check(1).EmpName
''        OSIFRec(1).Amt = Check(1).NetPay
''        OSIFRec(1).Src = 1
''        Put OHandle, OSNumRec, OSIFRec(1)
''        OSNumRec = OSNumRec + 1
''      End If
''      '** Update DISK files
''      Put CHandle, cnt, Check(1)
''      Put THandle, cnt, TransRec(1)
''      Put HHandle, NextHistRec, TransRec(1)
''      Put EHandle2, cnt, EmpRec2(1)
''      Put EHandle3, cnt, EmpRec3(1)
''      GoSub Save2VoidChecks  '6/17/04
''      NextHistRec& = NextHistRec& + 1
''    End If
''  Next
''
''  Close EHandle2
''  Close EHandle3
''  Close CHandle
''  Close THandle
''  Close HHandle
''  Close TVHandle
''
''  OpenTransHistFile HHandle
''  Get HHandle, NextHistRec& - 1, TransRec(1)
''  Close HHandle
''
''  If DoOSChkFlag Then
''    Close OHandle
''  End If
''
''  ReDim GLRec(1) As GLTransRecType
''  ReDim GLIFRec14(1) As GLIFDataType14
''  GLIFRecLen = Len(GLIFRec14(1))
''  GLRecLen = Len(GLRec(1))
''
''  'process gl transfer file
''
''  PRIF$ = "PRIF.DAT"
''  NextRec& = 1
''
''  NumOfRecs = FileSize("TEMPIF.DAT") \ GLIFRecLen
''
'''GLHandle closed immediately after retrieving GLSetupRec(1)
''  IHandle = FreeFile
''  Open "TEMPIF.DAT" For Random Shared As IHandle Len = GLIFRecLen
''
''  If Exist("PRIF.DAT") Then 'added 6/27/03 if this file is not deleted first
''  'then it causes the old file to be repeated when posting to GL
''    KillFile "PRIF.DAT"
''  End If
''
''  OHandle = FreeFile
''  Open PRIF$ For Random Shared As OHandle Len = GLRecLen
''  For cnt = 1 To NumOfRecs
''    Get IHandle, cnt, GLIFRec14(1)
''    GLRec(1).AcctNum = FmtAcct$(GLIFRec14(1).TranAcct, GLSetupRec(1).FundLen, GLSetupRec(1).AcctLen, GLSetupRec(1).DetLen)
''
''    'unrem
''    GLRec(1).TRDATE = GLDate
''    GLRec(1).Desc = GLIFRec14(1).TranDesc
''    '9/20/04 add the new Desc field here (see Paula)
''    GLRec(1).CrAmt = GLIFRec14(1).CrAmt
''    GLRec(1).DrAmt = GLIFRec14(1).DrAmt
''
''    GLRec(1).Src = FixDateSuffix(GLIFRec14(1).Source)
''    Put OHandle, NextRec, GLRec(1)
''    NextRec& = NextRec& + 1
''  Next
''
''  Close IHandle
''  Close OHandle
''  Post2GL PRIF$, PSysRec(), BadAccts   'unrem
''
'''  If BadAccts <> 0 Then
'''    'there was a posting to gl error
'''  Else
'''    KillFile "PRIF.DAT"
'''  End If
''
''  If Exist("TEMPIF.DAT") Then
''    KillFile "LASTIF.DAT"
''    Name "TEMPIF.DAT" As "LASTIF.DAT"
''  End If
''
''  'done with gl transfer file
''  Exit Sub
'
''Save2VoidChecks:
''  For VCnt = 1 To NumOfTempVoids
''    Get TVHandle, VCnt, TempVoid
''    If QPTrim$(TempVoid.EmpNum) = QPTrim$(EmpRec2(1).EmpNo) Then
''      If TempVoid.CheckNum = Check(1).CheckNum Then
''        TempVoid.TransRec = NextHistRec&
''        Put TVHandle, VCnt, TempVoid
''      End If
''    End If
''  Next VCnt
''
''  Return
''
''End Sub
'
'Function GetCitiDirFolder()
'  ReDim SysRec(1) As RegDSysFileRecType
'  Dim SysFileHandle As Integer
'  Dim TmpChr As String
'  Dim TmpDir As String
'
'  On Local Error Resume Next
'  OpenSysFile SysFileHandle
'  Get SysFileHandle, 1, SysRec(1)
'  Close SysFileHandle
'  'In the System Interface screen you cannot access the needed
'  'GL list if nothing is saved...this is an effort at allowing a gl
'  'search to occur if there is at least an entry in the Citipak
'  'field
'  'this function is also used to access gl files that are located
'  'only in the Citipak directory
''  TmpDir = QPTrim$(SysRec(1).CITIDIR)
'  TmpDir = CurrCitiPath
'  If Len(TmpDir) = 0 Then
'    GoTo PathOK
'  End If
'
'  TmpChr = Right$(TmpDir, 1)
'  If TmpChr = ":" Then
'    GetCitiDirFolder = TmpDir
'    GoTo PathOK
'  ElseIf TmpChr <> "\" Then
'    GetCitiDirFolder = TmpDir + "\"
'    GoTo PathOK
'  Else
'    GetCitiDirFolder = TmpDir
'  End If
'
'PathOK:
'  On Error GoTo 0
'End Function
'
'Function FixDateSuffix(Source As String)
'
'  Dim PDRRec As PeriodDefaultRecType
'  Dim PHandle As Integer
'  Dim EndDate As String
'  Dim EDLen As String
'  Dim TwoDigitYear As String
'  Dim SourceLen As Integer
'  'used in post transactions
'  OpenPPDefaultFile PHandle
'  Get PHandle, 1, PDRRec
'  Close PHandle
'  SourceLen = Len(Source)
'
'  EndDate = MakeRegDate(PDRRec.PEREND)
'  EDLen = Len(QPTrim$(EndDate))
'  TwoDigitYear = Mid(EndDate, EDLen - 1, 2)
'  FixDateSuffix = Mid(Source, 1, SourceLen - 2) + TwoDigitYear
'
'End Function
'
'Public Function SSNCheck(SSN As String) As Boolean
'  Dim SSN1 As String
'  Dim SSN2 As String
'  Dim SSN3 As String
'  'this function scans a social security number to make sure that
'  'it complies with the governmet's requirements
'  SSNCheck = False
'  SSN1 = Mid(SSN, 1, 3)
'  SSN2 = Mid(SSN, 5, 2)
'  SSN3 = Mid(SSN, 8, 4)
'
'  If SSN1 = "00" Or SSN2 = "000" Or SSN3 = "0000" Then
'    SSNCheck = True 'it's bad
'    Exit Function
'  End If
'
'  If Val(SSN1) = 666 Or Val(SSN1) = 680 Or (Val(SSN1) > 728 And Val(SSN1) < 750) Or Val(SSN1) = 764 Or Val(SSN1) >= 765 Then
'    SSNCheck = True 'can't include these numbers in these places
'  End If
'End Function
'
'Public Function FilesROK(frm As Form, InFileNames() As String, OutFileNames() As String, ThisMany As Integer) As Boolean
''  Dim NextName As Integer
''  Dim x As Integer
''  'this function scans for files necessary to run a particular part
''  'of the program and looks in the PRData folder for them...if they
''  'are missing then a warning screen pops up telling the user what
''  'the problem is and how to fix it (located in frmWarnFilesMissing)
''  FilesROK = True
''  NextName = 1
''  For x = 1 To ThisMany 'for loop takes incoming files needing checking
''  'and looks in PRData for them...if they are missing they are added
''  'to OutFileNames and if they are OK then they are skipped
''    If Not Exist(InFileNames(x)) Then
''      OutFileNames(NextName) = InFileNames(x)
''      NextName = NextName + 1
''      FilesROK = False
''    End If
''  Next x
''  If FilesROK = False Then
''    frmWarnFilesMissing.Show vbModal, frm
''    For x = 1 To ThisMany
''      InFileNames(x) = ""
''      OutFileNames(x) = ""
''    Next x
''  End If
'End Function
'
'
'Public Sub GetAcctStruct(CitipakPath$, GLFundLen%, GLAcctLen%, GLDetLen%)
'  Dim SetUpRecLen As Integer, SetupFile As Integer
'  ReDim GLSetupRec(1) As GLSetupRecType
'  'this sub determines the lengths of each piece of the gl number...
'  '(ie. 12-345-6789 breaks down to GLFundLen = 2, GLAcctLen = 3
'  'and GLDetLen (Dept) = 4)...this data is used in validating
'  'GL numbers before they are saved
''  StartPath = StartPath
'  SetUpRecLen = Len(GLSetupRec(1))
'  If Exist(QPTrim$(CitipakPath) + "GLSETUP.DAT") Then
'    SetupFile = FreeFile
'    Open CitipakPath + "GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
'  ElseIf Exist(QPTrim$(CitipakPath) + "\GLSETUP.DAT") Then
'    SetupFile = FreeFile
'    Open QPTrim$(CitipakPath) + "\GLSETUP.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
'  Else
'    Exit Sub
'  End If
'  Get SetupFile, 1, GLSetupRec(1)
'  Close SetupFile
'  GLFundLen = GLSetupRec(1).FundLen
'  GLAcctLen = GLSetupRec(1).AcctLen
'  GLDetLen = GLSetupRec(1).DetLen
'  Erase GLSetupRec
'End Sub
''Sub RPTSetupPRN(RPTNum, Handle)
''  Dim RPTPitch As Integer
''  Dim PrinterSetUpFile As Integer
''  Dim PrntType As PRNSetupRecType
''  Dim x As Integer
''  Dim PHandle As Integer
''  Dim DefPrinter As String
''  Dim PrnDef As String
''  Dim LineLen As Integer
''  Dim TextLine$
''  Dim Y As Integer
''  Dim z As Integer
''  Dim NextCommaPOS As Integer
''  Dim CodeStartPOS As Integer
''  Dim Codeline1$
''  Dim Codeline2$
''
''  'this sub coordinates the printing procedure so that any
''  'pitch data saved in the Printer setup screen for a
''  'particular report gets sent to the printer
''  For z = 1 To 10 'clear all existing codes
''    ToPrint1(z) = 0
''    ToPrint2(z) = 0
''  Next z
''  OpenPrinterSetupFile PrinterSetUpFile
''  Get PrinterSetUpFile, 1, PrntType
''  Close PrinterSetUpFile
''  DefPrinter = QPTrim$(PrntType.Printer)
''  'if a pitch isn't saved for this print job then by
''  'default the pitch becomes 10
''
''  If RPTNum = 123 Then GoTo SkipThis '123 is an arbitrary
''  'number used to signify the end of a report that tells
''  'this program to look for the reset codes
''
''  RPTPitch = PrntType.RPT(RPTNum) 'pitch is specified in the
''  'printer setup screen
''
''SkipThis:
''
''  GoSub GetPrinterCodes
''  If Len(Codeline1) Then 'CodeLine1 represents the reset codes
''  'because in the prprndf.dat file the reset codes come
''  'before the pitch codes
''  'at this point the proper codes have been determined and
''  'the select statement tells the printer which codes to use
''    Select Case Y
''      Case 1:
''        Print #Handle, Chr(ToPrint1(1));
''      Case 2:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2));
''      Case 3:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3));
''      Case 4:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4));
''      Case 5:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5));
''      Case 6:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6));
''      Case 7:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7));
''      Case 8:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8));
''      Case 9:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9));
''      Case 10:
''        Print #Handle, Chr(ToPrint1(1)); Chr(ToPrint1(2)); Chr(ToPrint1(3)); Chr(ToPrint1(4)); Chr(ToPrint1(5)); Chr(ToPrint1(6)); Chr(ToPrint1(7)); Chr(ToPrint1(8)); Chr(ToPrint1(9)); Chr(ToPrint1(10));
''      Case Else:
''    End Select
''  ElseIf Len(Codeline2) Then 'CodeLine2 represents the pitch codes
''    Select Case Y 'Y = the number of codes
''      Case 1:
''        Print #Handle, Chr(ToPrint2(1));
''      Case 2:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2));
''      Case 3:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3));
''      Case 4:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4));
''      Case 5:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5));
''      Case 6:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6));
''      Case 7:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7));
''      Case 8:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8));
''      Case 9:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9));
''      Case 10:
''        Print #Handle, Chr(ToPrint2(1)); Chr(ToPrint2(2)); Chr(ToPrint2(3)); Chr(ToPrint2(4)); Chr(ToPrint2(5)); Chr(ToPrint2(6)); Chr(ToPrint2(7)); Chr(ToPrint2(8)); Chr(ToPrint2(9)); Chr(ToPrint2(10));
''      Case Else:
''    End Select
''  End If
''
''  Exit Sub
''
''GetPrinterCodes:
''  PHandle = FreeFile
''  Open "PRData\Prprndf.dat" For Input As #PHandle  ' Open file.
''  Line Input #PHandle, TextLine   ' Read first line into TextLine.
''   'the second line is where individual printers start their codes
''   NextCommaPOS = 1
''
''   Do While Not EOF(PHandle) And NextCommaPOS <> 0  ' Loop until end of file.
''     Line Input #PHandle, TextLine   ' Read next line into Textline.
''     If TextLine = "@" + DefPrinter$ Then 'locate the default printer
''
''         If EOF(PHandle) Then Exit Do 'if for some reason we get to the end of the file
''         'then exit
''         If RPTNum = 123 Then '123 tells this code that we want the
''         'reset codes
''           Line Input #PHandle, TextLine 'read next line which by convention
''           'will always be the reset code line
''             LineLen = Len(TextLine)
''             Codeline1 = Mid(TextLine, 11, LineLen) 'by convention
''             '11 is where the first reset code begins in this line
''             CodeStartPOS = 1
''             Y = 1
''             Do
''               NextCommaPOS = InStr(CodeStartPOS, Codeline1, ",") 'look for comma
''               If NextCommaPOS = 0 Then 'if comma pos = 0 then we have no more commas
''                 LineLen = Len(Codeline1)
''                 ToPrint1(Y) = CInt(Mid(Codeline1, CodeStartPOS, 3))
''                 Exit Do 'we're at the end so exit loop
''               End If
''               ToPrint1(Y) = CInt(Mid(Codeline1, CodeStartPOS, 3)) 'look for a comma
''               CodeStartPOS = NextCommaPOS + 1 'start just behind the last comma
''               Y = Y + 1 ' advance y until no more commas found
''             Loop Until NextCommaPOS = 0
''             GoTo XIsOne 'jump to outer loop
''         End If
''       Do
''         If TextLine = "" Then 'added 6/25/2004
''           Unload FrmShowPctComp
''           MsgBox "No printer pitch codes could be found. Check the 'Printer Setup' screen settings. Defaulting to pitch 10."
''           Close PHandle
''           Exit Sub
''         End If 'added 6/25/2004 ^^^
''         Line Input #PHandle, TextLine 'look for pitch codes
''         If Mid(TextLine, 1, 2) = RPTPitch Then 'we found the proper pitch
''           LineLen = Len(TextLine)
''           Codeline2 = Mid(TextLine, 11, LineLen) 'read this line into
''           'Codeline2
''           Y = 1
''           CodeStartPOS = 1
''           Do
''             NextCommaPOS = InStr(CodeStartPOS, Codeline2, ",")
''             If NextCommaPOS = 0 Then 'no more commas in line
''               LineLen = Len(Codeline2)
''               ToPrint2(Y) = CInt(Mid(Codeline2, CodeStartPOS, 3))
''               Exit Do
''             End If
''             ToPrint2(Y) = CInt(Mid(Codeline2, CodeStartPOS, 3)) 'keep looking
''             'for commas
''             CodeStartPOS = NextCommaPOS + 1
''             Y = Y + 1
''           Loop Until NextCommaPOS = 0
''
''           Exit Do
''         End If
''        Loop Until NextCommaPOS = 0
''XIsOne:
'''     If TextLineCnt = 0 Then
'''        Unload FrmShowPctComp
'''        MsgBox "No printer pitch codes could be found. Check the 'Printer Setup' screen settings."
'''        Close PHandle
'''        Exit Sub
'''      End If
''
''     End If 'ends if TextLine = @ + DefPrinter
''   Loop
''   Close #PHandle   ' Close file.
''   Return
''
''
''End Sub
'
''Public Function AddDashesToGLNumber(ByVal GLNum$, Fund As Integer, Dept As Integer, Detail As Integer)
''  Dim NewGLNum As String
''
''  If Mid(GLNum, Fund + 1, 1) <> "-" And Mid(GLNum, Fund + Dept + 2, 1) <> "-" Then
''    NewGLNum = Mid(GLNum, 1, Fund) + "-" + Mid(GLNum, Fund + 1, Dept) + "-" + Mid(GLNum, Fund + Dept + 1, Detail)
''    AddDashesToGLNumber = NewGLNum
''  Else
''    AddDashesToGLNumber = GLNum
''  End If
''
''End Function
'Public Function AddDashesToGLNumber(ByVal GLNum$, Fund As Integer, Dept As Integer, Detail As Integer)
'  Dim NewGLNum As String
'
'  GLNum$ = ReplaceString(GLNum$, "-", "")
'  NewGLNum = Mid(GLNum$, 1, Fund) + "-"
'  NewGLNum = NewGLNum + Mid(GLNum$, Fund + 1, Dept) + "-"
'  NewGLNum = NewGLNum + Mid(GLNum$, Fund + Dept + 1, Detail)
'  AddDashesToGLNumber = NewGLNum
'
'End Function
'
'Public Sub PostVoidCheckData()
'  '----added 6/17/04-----------
'  'this sub is used when payroll is posted...not when
'  'a check is voided
'  Dim TempVoid As VoidCheckType
'  Dim TVHandle As Integer
'  Dim NumOfTempVoids As Integer
'  Dim VoidRec As VoidCheckType
'  Dim VPHandle As Integer 'Void Post
'  Dim NumOfVoidPosts As Double
'  Dim VCnt As Integer
'
'  OpenTempVoidFile TVHandle
'  NumOfTempVoids = LOF(TVHandle) / Len(TempVoid)
'
'  OpenVoidChkPostFile VPHandle
'  NumOfVoidPosts = LOF(VPHandle) / Len(VoidRec)
'
'  For VCnt = 1 To NumOfTempVoids
'    Get TVHandle, VCnt, TempVoid
''      TempVoid.CheckNum = TempVoid.CheckNum
''      TempVoid.PPEAmt = TempVoid.PPEAmt
''      TempVoid.PPEGL = TempVoid.PPEGL
'      VoidRec = TempVoid
'      NumOfVoidPosts = NumOfVoidPosts + 1
'    Put VPHandle, NumOfVoidPosts, VoidRec
'  Next VCnt
'
'  Close TVHandle
'  Close VPHandle
'
'End Sub
'
'Public Function PostVoidChkToGL(CHKDATE As Integer, EmpNum$, ChkNum As Long) As Boolean
'  'this function is used at the time a check is voided
''  Dim SysDir$, AcctFileName$, TransFileName$
''  Dim AcctIndexName$
''  Dim Tran2Post As GLTransRecType        'Dim a buffer for the edit file
''  Dim TrRecLen As Long
''  Dim File2Post As Integer
''  Dim Num2POst As Long
''  Dim TransFileNum As Integer
''  Dim NumAccts As Long
''  Dim AcctFileNum As Integer
''  Dim Acct As GLAcctRecType
''  Dim AcctRecLen As Long
''  Dim cnt As Long, Prev&
''  Dim Posted As Long, NumTrans&
''  Dim TransPosted As Long
''  Dim VoidRec As VoidCheckType
''  Dim VHandle As Integer
''  Dim VoidCnt As Double
''  Dim x As Double
''  Dim RecdNum As Long
''  Dim AcctIdx As GLAcctIndexType
''  Dim Trans As GLTransRecType
'''  Dim ErrCnt As Integer
''  Dim DedCnt As Integer
''  Dim TransGL$
''  Dim TransCAmt As Double
''  Dim TransDAmt As Double
''  Dim FundLen As Integer
''  Dim DeptLen As Integer
''  Dim DetLen As Integer
''  Dim SysRec As RegDSysFileRecType
''  Dim SHandle As Integer
''  Dim Source$
''  Dim ThisDate As Integer
''  Dim ThisStrDate$
''  Dim TransDesc$
''  Dim ThisPRFund$
''  Dim Y As Integer
''  Dim PRCnt As Integer
''
''  PRCnt = 0
''  OpenSysFile SHandle
''  Get SHandle, 1, SysRec
''  Close SHandle
''
'''  Call GetAcctStruct(QPTrim$(SysRec.CITIDIR), FundLen, DeptLen, DetLen)
''  Call GetAcctStruct(CurrCitiPath, FundLen, DeptLen, DetLen)
''  Source = "VP"
''  ThisDate = Date2Num(Date)
''  ThisStrDate = MakeRegDate(ThisDate)
''  Source = Source + Mid(ThisStrDate, 1, 2) + Mid(ThisStrDate, 4, 2) + Mid(ThisStrDate, 9, 2)
''  PostVoidChkToGL = False
''
''  OpenVoidChkPostFile VHandle
''  VoidCnt = LOF(VHandle) / Len(VoidRec)
''
''  OpenGLAcctFile AcctFileNum 'GLACCT.DAT
''  NumAccts = LOF(AcctFileNum) / Len(Acct)
''
''  OpenGLTransFile TransFileNum 'GLTRANS.DAT
''  NumTrans& = LOF(TransFileNum) / Len(Tran2Post)
''
''  ReDim ErrAcct(1 To 1) As String
''  ReDim ErrType(1 To 1) As String
''  ReDim ErrAmt(1 To 1) As Double
''  ReDim PRNetTotals(1 To 1) As Double
''  ReDim PRGL(1 To 1) As String
''  ErrEmpNum = ""
''  '---------10.05.04----------\/--\/
''  'added this code because when PRNet added up to be negative
''  'then the post to GL could subtract the negative amount from
''  'the debit and the credit columns because of the manner in
''  'which the positive PRNet amounts and the negative PRNet
''  'amounts were read in the payroll processing
''  'this code goes ahead and figures the PRNet balances for each
''  'PRNet GL number and posts them first before sorting through the
''  'rest of the numbers that don't have negative amounts
''
''  For x = 1 To VoidCnt
''    Get VHandle, x, VoidRec
''      If VoidRec.CheckNum = Chknum Then
''        If VoidRec.CheckDate = ChkDate Then
''          If QPTrim$(VoidRec.EmpNum) = QPTrim$(EmpNum$) Then
''            GlobalCheckNum$ = CStr(VoidRec.CheckNum)
''            ErrEmpNum = QPTrim$(VoidRec.EmpNum)
''            If VoidRec.PRNet = 0 Then GoTo TryThis 'SkipPRNet(changed 5/15/06)
''            VoidRec.PRNetGL = AddDashesToGLNumber(VoidRec.PRNetGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.PRNetGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.PRNetGL)
''              TransCAmt = OldRound#(VoidRec.PRNet)
''              If PRCnt = 0 Then
''                PRCnt = PRCnt + 1
''                PRGL(PRCnt) = TransGL$
''                PRNetTotals(PRCnt) = TransCAmt
''              Else
''                For Y = 1 To PRCnt
''                  If PRGL(Y) = TransGL$ Then
''                    PRNetTotals(Y) = OldRound(PRNetTotals(Y) + TransCAmt)
''                    Exit For
''                  End If
''                Next Y
''                If Y > PRCnt Then
''                  PRCnt = PRCnt + 1
''                  ReDim Preserve PRGL(1 To PRCnt) As String
''                  ReDim Preserve PRNetTotals(1 To PRCnt) As Double
''                  PRGL(PRCnt) = TransGL$
''                  PRNetTotals(PRCnt) = TransCAmt
''                End If
''              End If
''TryThis: 'added 5/15/06
''              If x < VoidCnt Then
''                Get VHandle, x + 1, VoidRec
''                If VoidRec.CheckNum = Chknum Then
''                  GoTo SkipPRNet
''                Else
''                  GoTo PostIt
''                End If
''              ElseIf x = VoidCnt Then
''PostIt:
''                For Y = 1 To PRCnt
''                  TransGL$ = PRGL(Y)
''                  RecdNum = FindAcct(AcctIndexName$, TransGL$)
''                  Get AcctFileNum, RecdNum, Acct
''                    If RecdNum > 0 Then
''                      TransCAmt = PRNetTotals(Y)
''                    End If
''                    If TransCAmt < 0 Then
''                      Acct.Typ = "L"
''                      TransDAmt = TransCAmt
''                      TransCAmt = 0
''                    Else
''                      TransCAmt = -TransCAmt
''                      TransDAmt = 0
''                    End If
''                    TransDesc = "VPPRNET"
''                    GoSub PostData
''                Next Y
''             End If
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.PRNetGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.PRNet)
''            End If
''          End If
''        End If
''      End If
''SkipPRNet:
''  Next x
''  '--------------^^^^^^^^^^^^^^10.05.04
''
''  For x = 1 To VoidCnt
''    Get VHandle, x, VoidRec
''      If VoidRec.CheckNum = Chknum Then
''        If VoidRec.CheckDate = ChkDate Then
''          If QPTrim$(VoidRec.EmpNum) = QPTrim$(EmpNum$) Then
''            GlobalCheckNum$ = CStr(VoidRec.CheckNum)
''            ErrEmpNum = QPTrim$(VoidRec.EmpNum)
''            If VoidRec.FEDWHAmt = 0 Then GoTo Skip1
''            VoidRec.FEDWHGL = AddDashesToGLNumber(VoidRec.FEDWHGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.FEDWHGL)  'Verify account is in G/L
''            If RecdNum > 0 Then                  'if valid acct then proceed
''              Get AcctFileNum, RecdNum, Acct    'Get the account
''              TransGL$ = QPTrim$(VoidRec.FEDWHGL)
''              TransCAmt = OldRound#(VoidRec.FEDWHAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPFEDWH"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.FEDWHGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.FEDWHAmt)
''            End If
''
''Skip1:
''            If VoidRec.MEDMATCRAmt = 0 Then GoTo Skip2
''            VoidRec.MEDMATCRGL = AddDashesToGLNumber(VoidRec.MEDMATCRGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDMATCRGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.MEDMATCRGL)
''              TransCAmt = OldRound#(VoidRec.MEDMATCRAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPMEDMATCHCRE"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.MEDMATCRGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.MEDMATCRAmt)
''            End If
''
''Skip2:
''            If VoidRec.MEDMATDBAmt = 0 Then GoTo Skip3
''            VoidRec.MEDMATDBGL = AddDashesToGLNumber(VoidRec.MEDMATDBGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDMATDBGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.MEDMATDBGL)
''              TransCAmt = 0
''              TransDAmt = OldRound#(VoidRec.MEDMATDBAmt)
''              TransDAmt = -TransDAmt
''              TransDesc = "VPMEDMATCHDBT"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.MEDMATDBGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Debit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.MEDMATDBAmt)
''            End If
''
''Skip3:
''            If VoidRec.MEDWHAmt = 0 Then GoTo Skip4
''            VoidRec.MEDWHGL = AddDashesToGLNumber(VoidRec.MEDWHGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.MEDWHGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.MEDWHGL)
''              TransCAmt = OldRound#(VoidRec.MEDWHAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPMEDWH"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.MEDWHGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.MEDWHAmt)
''            End If
''
''Skip4:
'''            If VoidRec.PRNet = 0 Then GoTo Skip5
'''            VoidRec.PRNetGL = AddDashesToGLNumber(VoidRec.PRNetGL, FundLen, DeptLen, DetLen)
'''            RecdNum = FindAcct(AcctIndexName$, VoidRec.PRNetGL)
'''            If RecdNum > 0 Then
'''              Get AcctFileNum, RecdNum, Acct
'''              TransGL$ = QPTrim$(VoidRec.PRNetGL)
'''              TransCAmt = OldRound#(VoidRec.PRNet)
'''              If TransCAmt < 0 Then 'And VoidRec.Type = "C" Or VoidRec.Type = "I" Then 'when dealing with central depository
'''              'the PRNet can sometimes be a negative...this means we have to
'''              'reverse the thinking on PRNet at this juncture
'''                Acct.Typ = "L"
'''                TransDAmt = TransCAmt
'''                TransCAmt = 0
'''              Else
'''                TransCAmt = -TransCAmt
'''                TransDAmt = 0
'''              End If
'''              TransDesc = "VPPRNET"
'''              GoSub PostData
'''            Else
'''              ErrCnt = ErrCnt + 1
'''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
'''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.PRNetGL)
'''              ReDim Preserve ErrType(1 To ErrCnt) As String
'''              ErrType(ErrCnt) = "Credit"
'''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
'''              ErrAmt(ErrCnt) = OldRound#(VoidRec.PRNet)
'''            End If
''
''Skip5:
''            If VoidRec.RETMATCRAmt = 0 Then GoTo Skip6
''            VoidRec.RETMATCRGL = AddDashesToGLNumber(VoidRec.RETMATCRGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETMATCRGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.RETMATCRGL)
''              TransCAmt = OldRound#(VoidRec.RETMATCRAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPRETMATCHCRE"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.RETMATCRGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.RETMATCRAmt)
''            End If
''
''Skip6:
''            If VoidRec.RETMATDBAmt = 0 Then GoTo Skip7
''            VoidRec.RETMATDBGL = AddDashesToGLNumber(VoidRec.RETMATDBGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETMATDBGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.RETMATDBGL)
''              TransCAmt = 0
''              TransDAmt = OldRound#(VoidRec.RETMATDBAmt)
''              TransDAmt = -TransDAmt
''              TransDesc = "VPRETMATCHDBT"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.RETMATDBGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Debit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.RETMATDBAmt)
''            End If
''
''Skip7:
''            If VoidRec.RETWHAmt = 0 Then GoTo Skip8
''            VoidRec.RETWHGL = AddDashesToGLNumber(VoidRec.RETWHGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.RETWHGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.RETWHGL)
''              TransCAmt = OldRound#(VoidRec.RETWHAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPRETWH"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.RETWHGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.RETWHAmt)
''            End If
''
''Skip8:
''            If VoidRec.SOCMATCRAmt = 0 Then GoTo Skip9
''            VoidRec.SOCMATCRGL = AddDashesToGLNumber(VoidRec.SOCMATCRGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCMATCRGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.SOCMATCRGL)
''              TransCAmt = OldRound#(VoidRec.SOCMATCRAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPSOCMATCHCRE"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.SOCMATCRGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.SOCMATCRAmt)
''            End If
''
''Skip9:
''            If VoidRec.SOCMATDBAmt = 0 Then GoTo Skip10
''            VoidRec.SOCMATDBGL = AddDashesToGLNumber(VoidRec.SOCMATDBGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCMATDBGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.SOCMATDBGL)
''              TransCAmt = 0
''              TransDAmt = OldRound#(VoidRec.SOCMATDBAmt)
''              TransDAmt = -TransDAmt
''              TransDesc = "VPSOCMATCHDBT"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.SOCMATDBGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Debit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.SOCMATDBAmt)
''            End If
''
''Skip10:
''            If VoidRec.SOCWHAmt = 0 Then GoTo Skip11
''            VoidRec.SOCWHGL = AddDashesToGLNumber(VoidRec.SOCWHGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.SOCWHGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.SOCWHGL)
''              TransCAmt = OldRound#(VoidRec.SOCWHAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPSOCWH"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.SOCWHGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.SOCWHAmt)
''            End If
''
''Skip11:
''            If VoidRec.STAWHAmt = 0 Then GoTo Skip12
''            VoidRec.STAWHGL = AddDashesToGLNumber(VoidRec.STAWHGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.STAWHGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.STAWHGL)
''              TransCAmt = OldRound#(VoidRec.STAWHAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPSTAWH"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.STAWHGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.STAWHAmt)
''            End If
''
''Skip12:
''            If VoidRec.WagesAmt = 0 Then GoTo Skip13
''            VoidRec.WagesGL = AddDashesToGLNumber(VoidRec.WagesGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.WagesGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.WagesGL)
''              TransCAmt = 0
''              TransDAmt = OldRound#(VoidRec.WagesAmt)
''              TransDAmt = -TransDAmt
''              TransDesc = "VPWAGES"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.WagesGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Debit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.WagesAmt)
''            End If
''
''Skip13:
''            If VoidRec.PPEAmt = 0 Then GoTo Skip14
''            VoidRec.PPEGL = AddDashesToGLNumber(VoidRec.PPEGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.PPEGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.PPEGL)
''              TransCAmt = 0
''              TransDAmt = OldRound#(VoidRec.PPEAmt)
''              TransDAmt = -TransDAmt
''              TransDesc = "VPCENTPPE"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.PPEGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Debit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.PPEAmt)
''            End If
''
''Skip14:
''            If VoidRec.PPETotAmt = 0 Then GoTo Skip15
''            VoidRec.PPETotGL = AddDashesToGLNumber(VoidRec.PPETotGL, FundLen, DeptLen, DetLen)
''            RecdNum = FindAcct(AcctIndexName$, VoidRec.PPETotGL)
''            If RecdNum > 0 Then
''              Get AcctFileNum, RecdNum, Acct
''              TransGL$ = QPTrim$(VoidRec.PPETotGL)
''              TransCAmt = OldRound#(VoidRec.PPETotAmt)
''              TransCAmt = -TransCAmt
''              TransDAmt = 0
''              TransDesc = "VPCENTPPETOTAL"
''              GoSub PostData
''            Else
''              ErrCnt = ErrCnt + 1
''              ReDim Preserve ErrAcct(1 To ErrCnt) As String
''              ErrAcct(ErrCnt) = QPTrim$(VoidRec.PPETotGL)
''              ReDim Preserve ErrType(1 To ErrCnt) As String
''              ErrType(ErrCnt) = "Credit"
''              ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''              ErrAmt(ErrCnt) = OldRound#(VoidRec.PPETotAmt)
''            End If
''
''Skip15:
''            For DedCnt = 1 To 50
''              If VoidRec.DedData(DedCnt).DAmt <> 0 Then
''                VoidRec.DedData(DedCnt).DedGLNum = AddDashesToGLNumber(VoidRec.DedData(DedCnt).DedGLNum, FundLen, DeptLen, DetLen)
''                RecdNum = FindAcct(AcctIndexName$, VoidRec.DedData(DedCnt).DedGLNum)
''                If RecdNum > 0 Then
''                  Get AcctFileNum, RecdNum, Acct
''                  TransGL$ = QPTrim$(VoidRec.DedData(DedCnt).DedGLNum)
''                  TransCAmt = OldRound#(VoidRec.DedData(DedCnt).DAmt)
''                  TransCAmt = -TransCAmt
''                  TransDAmt = 0
''                  TransDesc = QPTrim$(VoidRec.DedData(DedCnt).DedDesc)
''                  GoSub PostData
''                Else
''                  ErrCnt = ErrCnt + 1
''                  ReDim Preserve ErrAcct(1 To ErrCnt) As String
''                  ErrAcct(ErrCnt) = QPTrim$(VoidRec.DedData(DedCnt).DedGLNum)
''                  ReDim Preserve ErrType(1 To ErrCnt) As String
''                  ErrType(ErrCnt) = "Credit"
''                  ReDim Preserve ErrAmt(1 To ErrCnt) As Double
''                  ErrAmt(ErrCnt) = OldRound#(VoidRec.DedData(DedCnt).DAmt)
''                End If
''              End If
''            Next DedCnt
''          End If
''        End If
''        VoidRec.VoidFlag = True
''      End If
'''    Put VHandle, X, VoidRec
''  Next x
''
''  If ErrCnt > 0 Then
''    frmPRVoidGLError.Show vbModal
''  End If
''
''  Close
''
''  Exit Function
''
''PostData:
''  PostVoidChkToGL = True
''  Select Case Acct.Typ
''    Case "A", "E"                 'asset, exp accts
''      Acct.Bal = OldRound#(Acct.Bal) + OldRound#(TransDAmt) - OldRound#(TransCAmt)
''      Put AcctFileNum, RecdNum, Acct
''
''    Case "L", "R"                 'liab, rev accts
''      Acct.Bal = OldRound#(Acct.Bal) + OldRound#(TransCAmt) - OldRound#(TransDAmt)
''      Put AcctFileNum, RecdNum, Acct
''  End Select
''
''  NumTrans& = NumTrans& + 1          'increment record pointer
''  Get TransFileNum, NumTrans&, Trans
''  Trans.AcctNum = TransGL$
''  Trans.TRDATE = frmVoidListOfChecks.WhichDate '02/01/05
'''  Trans.TRDATE = Date2Num(Date)
''  Trans.Desc = TransDesc
''  Trans.CrAmt = TransCAmt
''  Trans.DrAmt = TransDAmt
''  Trans.Ref = ""
''  Trans.Src = Source
''  Trans.NextTran = 0
''
''  Put TransFileNum, NumTrans&, Trans
''  '---------------------------------Start linking here
''  If Acct.FrstTran = 0 Then        'if first trans for this acct,
''    Acct.FrstTran = NumTrans&      'assign first & last pointers to
''    Acct.LastTran = NumTrans&      'this transaction
''    Put AcctFileNum, RecdNum, Acct
''  Else                             'otherwise
''    Prev& = Acct.LastTran             'remember the prev trans pointer,
''    Acct.LastTran = NumTrans&        'reset last trans to this trans
''    Put AcctFileNum, RecdNum, Acct
''                                   'In the trans file...
''    Get TransFileNum, Prev&, Trans    'Get the last transaction
''    Trans.NextTran = NumTrans&       'reset pointer to this trans
''    Put TransFileNum, Prev&, Trans
''  End If
''
''  Return
''
''
''   'LOCK AcctFileNum
'''   OpenGLAcctFile AcctFileNum 'GLACCT.DAT
'''   NumAccts = LOF(AcctFileNum) \ Len(Acct)
''
''   'LOCK TransFileNum
'''   OpenGLTransFile TransFileNum 'GLTRANS.DAT
'''   NumTrans& = LOF(TransFileNum) / Len(Tran2Post)
'''   For cnt = 1 To Num2POst                'Start processing transactions
'''     Get File2Post, cnt, Tran2Post
'''     RecdNum = FindAcct(AcctIndexName$, Tran2Post.AcctNum)  'Verify account is in G/L
'''     If RecdNum > 0 Then                  'if valid acct then proceed
'''       Get AcctFileNum, RecdNum, Acct    'Get the account
'''       'depending on account type, update running balance
'''       'Nick was updating MTD & YTD fields here also.
'''
'''       Select Case Acct.Typ
'''         Case "A", "E"                 'asset, exp accts
'''           Acct.Bal = OldRound#(Acct.Bal) + OldRound#(TransDebit) - OldRound#(TransCredit)
'''           Put AcctFileNum, RecdNum, Acct
'''
'''         Case "L", "R"                 'liab, rev accts
'''           Acct.Bal = OldRound#(Acct.Bal) + OldRound#(Tran2Post.CrAmt) - OldRound#(Tran2Post.DrAmt)
'''           Put AcctFileNum, RecdNum, Acct
'''
'''       End Select
'''       NumTrans& = NumTrans& + 1          'increment record pointer
'''       Get TransFileNum, NumTrans&, Trans
'''       Trans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
'''       Trans.TrDate = Tran2Post.TrDate
'''       Trans.Desc = Tran2Post.Desc
'''       Trans.CrAmt = Tran2Post.CrAmt
'''       Trans.DrAmt = Tran2Post.DrAmt
'''       Trans.Ref = "" 'Tran2Post.Ref
'''       Trans.Src = Tran2Post.Src
'''       Trans.NextTran = 0
'''
'''       Put TransFileNum, NumTrans&, Trans
'''
'''       Posted = Posted + 1
'''
'''       '---------------------------------Start linking here
'''       If Acct.FrstTran = 0 Then        'if first trans for this acct,
'''         Acct.FrstTran = NumTrans&      'assign first & last pointers to
'''         Acct.LastTran = NumTrans&      'this transaction
'''         Put AcctFileNum, RecdNum, Acct
'''       Else                             'otherwise
'''         Prev& = Acct.LastTran             'remember the prev trans pointer,
'''         Acct.LastTran = NumTrans&        'reset last trans to this trans
'''         Put AcctFileNum, RecdNum, Acct
'''                                        'In the trans file...
'''         Get TransFileNum, Prev&, Trans    'Get the last transaction
'''         Trans.NextTran = NumTrans&       'reset pointer to this trans
'''         Put TransFileNum, Prev&, Trans
'''       End If
'''       TransPosted = TransPosted + 1
'''     Else                                'Account NOT found!
'''       BadTrans = BadTrans + 1          'Pass info back to caller
'''                                        'how about an error log here.
'''     End If
'''
'''   Next
'''
'''Close
'''
'''Exit Function
''
'''was printing register and deleteing edit file here.
'''Now do this in module that called this sub
''
'End Function
'
'Public Function LastDayOfMonth(Month As String, Year As String) As String
'
'  Select Case Month
'    Case "04", "06", "09", "11"
'      LastDayOfMonth = "30"
'    Case "02"
'      Select Case Year
'        Case "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60", "64", "68", "72", "76", "80", "84", "88", "92", "96", "00"
'          LastDayOfMonth = "29"
'        Case Else
'          LastDayOfMonth = "28"
'      End Select
'    Case Else
'      LastDayOfMonth = "31"
'  End Select
'
'End Function
'
'Public Sub MakeTransInactive()
'  Dim TransRec As TransRecType
'  Dim THandle As Integer
'  Dim NumOfRecs As Integer
'  Dim x As Integer
'  '8/20 added progress bar
'  KillFile ("prdata\ChecksPrinted.opn") '10/3/03
'  'without this Killfile the user can resave defaults and
'  'then go directly to post with no warnings.
'  FrmShowPctComp.Label1 = "Clearing Former Payroll Defaults"
''  FrmShowPctComp.cmdCancel.Visible = False
'  FrmShowPctComp.Show
'  DoEvents
''  EnableCloseButton Me.hwnd, False
''  Me.cmdExit.Enabled = False
''  Me.cmdSave.Enabled = False
'  OpenTransWorkFile THandle
'  NumOfRecs = LOF(THandle) \ Len(TransRec)
'  For x = 1 To NumOfRecs
'    Get THandle, x, TransRec
'    TransRec.TActive = False
'    Put THandle, x, TransRec
'    FrmShowPctComp.ShowPctComp x, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
''      EnableCloseButton Me.hwnd, True
''      Me.cmdExit.Enabled = True
''      Me.cmdSave.Enabled = True
'      Exit Sub
'    End If
'  Next x
'  Close THandle
'  Unload FrmShowPctComp
'End Sub
'
Public Function FileExists(ByVal strFileName As String) As Boolean
  On Error Resume Next
'
  If (Len(Dir$(strFileName)) > 0) Then
    FileExists = True
  Else
    FileExists = False
  End If
  On Error GoTo 0
End Function
'
Public Function DirExists(ByVal strDirName As String) As Boolean
  On Error Resume Next
'
  Dim strFileName As String
'
  strFileName = strDirName & "\Nul"
'
  If (FileExists(strFileName)) Then
    DirExists = True
  Else
    DirExists = False
  End If
  On Error GoTo 0
End Function
'
'Public Sub UpdatePayRate(JobName$, PayType$, OTRate As Double, RegRate As Double, Freq As String, RecNum, NewEmp As Boolean)
'  Dim PayRate As PayRateType
'  Dim PHandle As Integer
'  Dim NumOfPayRate As Integer
'  Dim EmpRec As EmpData2Type
'  Dim EHandle As Integer
'  Dim NumOfEmpRecs As Integer
'  Dim x As Integer
'  Dim Y As Integer
'  Dim z As Integer
'  Dim EmpNum$
'  Dim EmpFName$
'  Dim EmpLName$
'  Dim Frequency$
'  Dim thisNum$
'
'  OpenEmpData2File EHandle
'  NumOfEmpRecs = LOF(EHandle) / Len(EmpRec)
'  If NumOfEmpRecs = 0 Then
'    Close
'    Exit Sub
'  End If
'
'  OpenPayRateFile PHandle
'  NumOfPayRate = LOF(PHandle) / Len(PayRate)
'  If NumOfPayRate = 0 Then GoTo FirstTimeThru
'
'  Get EHandle, RecNum, EmpRec
'  Close EHandle
'
'  If NewEmp = True Then
'    GoSub SaveNewEmp
'    GoTo GotIt
'  End If
'
'  If OTRate <> EmpRec.EMPORATE Or UCase(QPTrim$(Freq)) <> UCase(QPTrim$(EmpRec.EMPPFREQ)) Or RegRate <> EmpRec.EMPPRATE Or UCase(QPTrim$(PayType)) <> UCase(QPTrim$(EmpRec.EMPPTYPE)) Then
'    Get PHandle, RecNum, PayRate
'    For z = 1 To 30
'      If z = 30 And PayRate.RegPayRate(z) <> 0 Then
'        MsgBox "There are no more available pay rate record slots available for this employee."
'        Close PHandle
'        Exit Sub
'      End If
'      If PayRate.RegPayRate(z) = 0 Then
'        If z = 30 Then
'          MsgBox "This is the last available pay rate record for this employee."
'        End If
'        PayRate.EmpFName = QPTrim$(EmpRec.EmpFName)
'        PayRate.EmpLName = QPTrim$(EmpRec.EmpLName)
'        PayRate.EMPTDATE = EmpRec.EMPTDATE
'        PayRate.EMPHDATE = EmpRec.EMPHDATE
'        PayRate.OTPayRate(z) = OTRate
'        PayRate.EMPPFREQ(z) = QPTrim$(Freq)
'        PayRate.EMPPTYPE(z) = QPTrim$(PayType)
'        PayRate.PayChngDate(z) = Date2Num(Date)
'        PayRate.RegPayRate(z) = RegRate
'        PayRate.EmpRecNum = RecNum
'        PayRate.EMPJOB(z) = QPTrim$(JobName)
'        Put PHandle, RecNum, PayRate
'        GoTo GotIt
'      End If
'    Next z
'  End If
'
'GotIt:
'  Close PHandle
'  Call MakePayRateIndex
'
'  Exit Sub
'
'SaveNewEmp:
'  PayRate.EmpFName = QPTrim$(EmpRec.EmpFName)
'  PayRate.EmpLName = QPTrim$(EmpRec.EmpLName)
'  PayRate.EmpNo = QPTrim$(EmpRec.EmpNo)
'  PayRate.EMPTDATE = EmpRec.EMPTDATE
'  PayRate.EMPHDATE = EmpRec.EMPHDATE
'  PayRate.OTPayRate(1) = OTRate
'  PayRate.EMPPFREQ(1) = QPTrim$(Freq)
'  PayRate.EMPPTYPE(1) = QPTrim$(PayType)
'  PayRate.PayChngDate(1) = Date2Num(Date)
'  PayRate.RegPayRate(1) = RegRate
'  PayRate.EmpRecNum = RecNum
'  PayRate.EMPJOB(1) = QPTrim$(EmpRec.EMPJOB)
'  Put PHandle, RecNum, PayRate
'
'  Return
'
'FirstTimeThru:
'  For x = 1 To NumOfEmpRecs
'    Get EHandle, x, EmpRec
'    PayRate.EmpFName = QPTrim$(EmpRec.EmpFName)
'    PayRate.EmpLName = QPTrim$(EmpRec.EmpLName)
'    PayRate.EmpNo = QPTrim$(EmpRec.EmpNo)
'    PayRate.EMPTDATE = EmpRec.EMPTDATE
'    PayRate.EMPHDATE = EmpRec.EMPHDATE
'    PayRate.EMPPFREQ(1) = QPTrim$(EmpRec.EMPPFREQ)
'    PayRate.EMPPTYPE(1) = QPTrim$(EmpRec.EMPPTYPE)
'    PayRate.PayChngDate(1) = Date2Num(Date)
'    PayRate.OTPayRate(1) = EmpRec.EMPORATE
'    PayRate.RegPayRate(1) = EmpRec.EMPPRATE
'    PayRate.EMPJOB(1) = QPTrim$(EmpRec.EMPJOB)
'    PayRate.EmpRecNum = x
'    Put PHandle, x, PayRate
'  Next x
'
' Close EHandle
' Close PHandle
' Call MakePayRateIndex
'
'End Sub
'
'Public Sub MakePayRateIndex()
'  Dim PayRec As PayRateType
'  Dim PHandle As Integer
'  Dim x As Integer
'  Dim IdxRec As PayRateIndexType
'  Dim IdxNumRec As PayRateIdxNumType
'  Dim XHandle As Integer
'  Dim NumOfPayRecs As Integer
'  Dim HoldThis As PayRateIndexType
'  Dim HoldThisNum As PayRateIdxNumType
'  Dim Big$, BigNum$
'  Dim ThisName$, thisNum$
'  Dim Nextx As Integer
'  Dim Small$
'  Dim Thisx As Integer
'  Dim ThisRec As Integer
'
'  On Error GoTo ERRORSTUFF
'
'  OpenPayRateFile PHandle
'  NumOfPayRecs = LOF(PHandle) / Len(PayRec)
'  If NumOfPayRecs = 0 Then
'    Close
'    Exit Sub
'  End If
'
'  Big$ = ""
'  BigNum = "0"
'  ReDim TempPayIdx(1 To NumOfPayRecs) As PayRateIndexType
'  ReDim TempPayNumIdx(1 To NumOfPayRecs) As PayRateIdxNumType
'
'  For x = 1 To NumOfPayRecs
'    Get PHandle, x, PayRec
'    ThisName = QPTrim$(PayRec.EmpLName) + QPTrim$(PayRec.EmpFName)
'    thisNum = QPTrim$(PayRec.EmpNo)
'    TempPayIdx(x).PayRateRec = PayRec.EmpRecNum
'    TempPayIdx(x).EmpName = ThisName
'    TempPayNumIdx(x).PayRateRec = PayRec.EmpRecNum
'    TempPayNumIdx(x).EmpNum = thisNum
'    If ThisName > Big$ Then
'      Big$ = ThisName
'    End If
'    If Val(thisNum) > Val(BigNum$) Then
'      BigNum$ = thisNum
'    End If
'  Next x
'  Close PHandle
'
'  Big = Big + "Z"
'  Small = Big
'  Nextx = 1
'  Do
'    For x = Nextx To NumOfPayRecs
'      ThisName = QPTrim$(TempPayIdx(x).EmpName)
'      If ThisName < Small Then
'        Small = ThisName
'        Thisx = x
'      End If
'    Next x
'    HoldThis = TempPayIdx(Nextx)
'    TempPayIdx(Nextx) = TempPayIdx(Thisx)
'    TempPayIdx(Thisx) = HoldThis
'    If Nextx = NumOfPayRecs Then Exit Do
'    Nextx = Nextx + 1
'    Small = Big
'
'  Loop
'
'  KillFile "PRDATA\PAYRTIDX.DAT"
'
'  OpenPayRateIdxFile XHandle
'  For x = 1 To NumOfPayRecs
'    IdxRec = TempPayIdx(x)
'    Put XHandle, x, IdxRec
'  Next x
'
'  Close XHandle
'  '-----------Index numbers---------------
'  BigNum = CStr(Val(BigNum) + 1)
'  Small = BigNum
'  Nextx = 1
'  Do
'    For x = Nextx To NumOfPayRecs
'      thisNum = QPTrim$(TempPayNumIdx(x).EmpNum)
'      If Val(thisNum) < Val(Small) Then
'        Small = thisNum
'        Thisx = x
'      End If
'    Next x
'    HoldThisNum = TempPayNumIdx(Nextx)
'    TempPayNumIdx(Nextx) = TempPayNumIdx(Thisx)
'    TempPayNumIdx(Thisx) = HoldThisNum
'    If Nextx = NumOfPayRecs Then Exit Do
'    Nextx = Nextx + 1
'    Small = BigNum
'
'  Loop
'
'  KillFile "PRDATA\PYRTNUMIDX.DAT"
'
'  OpenPayRateNumIdxFile XHandle
'  For x = 1 To NumOfPayRecs
'    IdxNumRec = TempPayNumIdx(x)
'    IdxNumRec.EmpNum = IdxNumRec.EmpNum
'    Put XHandle, x, IdxNumRec
'  Next x
'
'  Close XHandle
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "PRCommon", "MakePayRateIndex", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'    Terminate
'
'End Sub
'
'Public Function FigurePayIncPct(NewPaySalOrHrly$, OldPaySalOrHrly$, PayTypeOld$, PayTypeNew$, OldPay As Double, NewPay As Double) As Double
'  Dim Hrs As Double
'  Dim OldAnnual As Double
'  Dim NewAnnual As Double
'  Dim PayFreq As Integer
'
'  FigurePayIncPct = 0
'
'  If OldPay = 0 Or NewPay = 0 Then
'    FigurePayIncPct = 0
'    Exit Function
'  End If
'
'  Select Case QPTrim$(UCase$(PayTypeOld$))
'    Case "WEEKLY"
'      PayFreq = 52
'      Hrs# = 40
'    Case "BI-WEEKLY"
'      PayFreq = 26
'      Hrs# = 80
'    Case "SEMI-MONTHLY"
'      PayFreq = 24
'      Hrs# = 86.66
'    Case "MONTHLY"
'      PayFreq = 12
'      Hrs# = 173.33
'    Case "QUARTERLY"
'      PayFreq = 4
'      Hrs# = 520
'    Case "SEMI-ANNUALLY"
'      PayFreq = 2
'      Hrs# = 1040
'    Case "ANNUALLY"
'      PayFreq = 1
'      Hrs# = 2080
'  End Select
'
'  If QPTrim$(UCase(OldPaySalOrHrly)) = "HOURLY" Then
'    OldAnnual = OldRound(Hrs * OldPay * PayFreq)
'  ElseIf QPTrim$(UCase(OldPaySalOrHrly)) = "SALARIED" Then
'    OldAnnual = OldRound(PayFreq * OldPay)
'  End If
'
'  Select Case QPTrim$(UCase$(PayTypeNew$))
'    Case "WEEKLY"
'      PayFreq = 52
'      Hrs# = 40
'    Case "BI-WEEKLY"
'      PayFreq = 26
'      Hrs# = 80
'    Case "SEMI-MONTHLY"
'      PayFreq = 24
'      Hrs# = 86.66
'    Case "MONTHLY"
'      PayFreq = 12
'      Hrs# = 173.33
'    Case "QUARTERLY"
'      PayFreq = 4
'      Hrs# = 520
'    Case "SEMI-ANNUALLY"
'      PayFreq = 2
'      Hrs# = 1040
'    Case "ANNUALLY"
'      PayFreq = 1
'      Hrs# = 2080
'  End Select
'
'  If QPTrim$(UCase(NewPaySalOrHrly)) = "HOURLY" Then
'    NewAnnual = OldRound(Hrs * NewPay * PayFreq)
'  ElseIf QPTrim$(UCase(NewPaySalOrHrly)) = "SALARIED" Then
'    NewAnnual = OldRound(PayFreq * NewPay)
'  End If
'
'  FigurePayIncPct = (NewAnnual - OldAnnual) / OldAnnual
'
'End Function
'
'Public Sub UpdatePayRateEscapeVrs(JobName$, NewPType$, NewORate#, NewPRate#, NewPFreq$, OldPType$, OldORate#, OldPRate#, OldPFreq$, ThisRec%)
'  Dim PayRate As PayRateType
'  Dim PHandle As Integer
'  Dim NumOfPayRate As Integer
'  Dim EmpRec As EmpData2Type
'  Dim EHandle As Integer
'  Dim NumOfEmpRecs As Integer
'  Dim x As Integer
'  Dim Y As Integer
'  Dim z As Integer
'  Dim EmpNum$
'  Dim EmpFName$
'  Dim EmpLName$
'  Dim Frequency$
'
'  OpenEmpData2File EHandle
'  NumOfEmpRecs = LOF(EHandle) / Len(EmpRec)
'  If NumOfEmpRecs = 0 Then
'    Close
'    Exit Sub
'  End If
'
'  OpenPayRateFile PHandle
'  NumOfPayRate = LOF(PHandle) / Len(PayRate)
'  If NumOfPayRate = 0 Then Exit Sub
'
'  Get EHandle, ThisRec, EmpRec
'  Close EHandle
'
'  If NewORate <> OldORate Or UCase(QPTrim$(NewPFreq)) <> UCase(QPTrim$(OldPFreq)) Or NewPRate <> OldPRate Or UCase(QPTrim$(NewPType)) <> UCase(QPTrim$(OldPType)) Then
'    Get PHandle, ThisRec, PayRate
'    For z = 1 To 30
'      If z = 30 And PayRate.RegPayRate(z) <> 0 Then
'        MsgBox "There are no more available pay rate record slots available for this employee."
'        Close PHandle
'        Exit Sub
'      End If
'      If PayRate.RegPayRate(z) = 0 Then
'        If z = 30 Then
'          MsgBox "This is the last available pay rate record for this employee."
'        End If
'        PayRate.EmpFName = QPTrim$(EmpRec.EmpFName)
'        PayRate.EmpLName = QPTrim$(EmpRec.EmpLName)
'        PayRate.EMPTDATE = EmpRec.EMPTDATE
'        PayRate.EMPHDATE = EmpRec.EMPHDATE
'        PayRate.OTPayRate(z) = NewORate
'        PayRate.EMPPFREQ(z) = QPTrim$(NewPFreq)
'        PayRate.EMPPTYPE(z) = QPTrim$(NewPType)
'        PayRate.PayChngDate(z) = Date2Num(Date)
'        PayRate.RegPayRate(z) = NewPRate
'        PayRate.EmpRecNum = ThisRec
'        PayRate.EMPJOB(z) = QPTrim$(JobName)
'        Put PHandle, ThisRec, PayRate
'        GoTo GotIt
'      End If
'    Next z
'  End If
'
'GotIt:
'  Close PHandle
'  Call MakePayRateIndex
'
'End Sub
'
''Public Sub GetTemp()
''  Dim Tempfile As Integer, lentemp As Integer
''  Dim PassTemp As CitiPassTempType
''
''  'lentemp = Len(Tempfile)
''  Tempfile = FreeFile
''  Open "c:\PassTemp.dat" For Random Shared As Tempfile ' Len = lentemp
''  Get Tempfile, 1, PassTemp
''  PWUser = QPTrim(PassTemp.UserName)
''  PWcnt = PassTemp.usernum
''  Close
''
''End Sub
'
''Public Sub SetToGo()
''  Dim Tempfile As Integer, lentemp As Integer
''  Dim PassTemp As CitiPassTempType
''
''  Tempfile = FreeFile
''  Open "c:\PassTemp.dat" For Random Shared As Tempfile ' Len = lentemp
''  PassTemp.usernum = PWcnt
''  PassTemp.UserName = PWUser
''  PassTemp.frommdl = 99
''  Put Tempfile, 1, PassTemp
''  Close
''End Sub
'
'Public Function CustHasMsg(RecNo&) As Boolean
'
'  Dim MsgRec As PRMessRecType
'  Dim MsgHandle As Integer, PRCustRecLen As Integer
'  Dim x As Integer, Y As Integer
'  Dim NumMsgRec As Long, MRec As Long
'
'  CustHasMsg = False
'  OpenEmpMessage MsgHandle
'  NumMsgRec& = LOF(MsgHandle) / Len(MsgRec)
'  If NumMsgRec& = 0 Then
'    Close MsgHandle
'    Exit Function
'  End If
'
'  If RecNo& > 0 Then
'    For x = 1 To NumMsgRec
'      Get MsgHandle, x, MsgRec
'      If MsgRec.EmpRec = RecNo Then
'        For Y = 1 To 15
'          If Len(QPTrim$(MsgRec.MessLine(Y).Msg)) > 0 Then
'            CustHasMsg = True
'            Exit For
'          End If
'        Next Y
'        Exit For
'      End If
'    Next x
'  End If
'
'End Function
'
'Public Function RemNulls$(Text As String)
'  Dim StrLen As Long
'  Dim Cnt As Long
'  Dim thischar As Integer
'  StrLen = Len(Text)
'  For Cnt = 1 To StrLen
'    thischar = Asc(Mid$(Text, Cnt, 1))
'    If thischar = 0 Then
'      Mid$(Text$, Cnt, 1) = " "
'    End If
'  Next
'  RemNulls$ = Text
'End Function

