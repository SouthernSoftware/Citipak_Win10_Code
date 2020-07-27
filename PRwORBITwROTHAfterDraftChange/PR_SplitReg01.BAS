Attribute VB_Name = "modSplit01"
Option Explicit
  Dim GLCreditTotal As Double
  Dim GLDebitTotal As Double
  Dim GLError As Integer
  Dim FundArray() As String '8/18/04
  Dim FundCount As Integer
  Dim FedTaxByFund() As Double '8/10/04
  Dim EICByFund() As Double '5/6/09
  Dim MedTaxByFund() As Double '8/10/04
  Dim MedMatByFund() As Double '8/10/04
  Dim SocTaxByFund() As Double '8/10/04
  Dim SocMatByFund() As Double '8/10/04
  Dim StaTaxByFund() As Double '8/10/04
  Dim RetTaxByFund() As Double '8/10/04
  Dim RetMatByFund() As Double '8/10/04
  Dim DedAmtByFund() As Double '8/10/04
  Dim TOTFEDTAX As Double '8/13/04
  Dim TOTMEDTAX As Double '8/13/04
  Dim TotEIC As Double '5/6/09
  Dim TOTSOCTAX As Double '8/13/04
  Dim TOTSTATAX As Double '8/13/04
  Dim TOTRetTax As Double '8/13/04
  Dim TOTMEDMat As Double '8/13/04
  Dim TOTSOCMat As Double '8/13/04
  Dim TOTRETMat As Double '8/13/04
  Dim TotDedAmt() As Double '8/13/04
  Dim FundLen As Integer
Sub PCPrintPayRegisterSG(PathCode As Integer)
  
  Dim TransRec(1) As TransRecType
  Dim EmpRec1(1) As EmpData1Type
  Dim PDR(1) As PeriodDefaultRecType
  Dim Unit(1) As UnitFileRecType

  ReDim DistbSumAccts(1 To 1) As DistWageRptType
  Dim SysRec(1) As RegDSysFileRecType
  Dim GLIFRec(1 To 1) As GLIFDataType12
  ReDim SplitDeds(1 To 1) As SplitDedsType
  ReDim EmpSpl(1 To 1) As SplitDedsType
  ReDim EmpAct(1 To 1) As DistWageRptType
  Dim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  Dim ErnCodes(1 To 3) As ErnCodeRecType
  Dim TotDeds#(1 To 50), PrintGLRpt As Boolean
  Dim TotErns(1 To 3) As Double
  
  Dim EDesc(1) As String * 21
  Dim EDAct(1) As String * 14
  Dim EDPct(1) As String * 11
  Dim EDRHrs(1) As String * 11
  Dim EDOHrs(1) As String * 11
  Dim EDRPay(1) As String * 11
  Dim EDOPay(1) As String * 11
  Dim EDEarn(1) As String * 11
  Dim EDGroP(1) As String * 11
  Dim EDSAmt(1) As String * 11
  Dim EDMAmt(1) As String * 11
  Dim EDRAmt(1) As String * 11
  Dim ENumb(1) As String * 14 'was ENum
  Dim EName(1) As String * 33

  Dim BRat(1) As String * 11
  Dim ORat(1) As String * 11
  Dim TaxFrn(1) As String * 11
  Dim Fill11(1) As String * 11
  Dim SCnt(1) As String * 11
  Dim HCnt(1) As String * 11
  Dim RHrs(1) As String * 11
  Dim VHrs(1) As String * 11
  Dim SHrs(1) As String * 11
  Dim HHrs(1) As String * 11
  Dim CHrs(1) As String * 11

  Dim THrs(1) As String * 11
  Dim PHrs(1) As String * 11
  Dim OTHrs(1) As String * 11
  Dim OTPaid(1) As String * 11
  Dim OTComp(1) As String * 11
  Dim RErnP(1) As String * 11
  Dim OErnP(1) As String * 11
  Dim GPayP(1) As String * 11
  Dim SSTaxP(1) As String * 11
  Dim MTaxP(1) As String * 11
  Dim FTaxP(1) As String * 11
  Dim STaxP(1) As String * 11
  Dim RetirP(1) As String * 11
  Dim NetPayP(1) As String * 11
  Dim Ded(1) As String * 11

  'added for EIC   6/07/94
  Dim EEicP(1) As String * 11
  Dim Ern(1) As String * 11
  Dim Pg(1) As String * 5
  Dim EMPLine(1) As String * 132
  Dim SplDed(1) As String * 7
  Dim SplDed2(1) As String * 9
  Dim Dash0(1) As String * 69
  Dim Dash(1) As String * 132
  Dim Dash2(1) As String * 124
  Dim Dash3(1) As String * 91
  Dim PPDefaultHandle As Integer
  Dim FileHandle As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim Emp1Handle As Integer
  Dim x As Integer, LincCnt As Integer, y As Integer
  Dim EmpIdxNNameHandle As Integer
  Dim NumIdxRec As NumbSortIdxType
  Dim FACnt As Integer, RetCode As Integer
  Dim Image0$, Image$, Image3$, Image4$, Image5$
  Dim GFedGross#, GStaGross#, GMedGross#, GSocGross#
  Dim GRetGross#, GTaxFrn#, TotDebit#, TotCredit#
  Dim ArraySize As Integer, EmpActLen As Integer
  Dim DistbSumAcctsLen As Integer, TransRecLen As Long
  Dim Emp1RecLen As Integer, IdxRecLen As Integer
  Dim IdxFileSize&, NumOfRecs As Long, SalCnt As Integer
  Dim HrlCnt As Integer, DLineCnt As Integer
  Dim LineCnt As Integer, NumOfWageAccts As Integer
  Dim MaxLines As Integer, DMaxLines As Integer
  Dim EPage As Integer, Page As Integer
  Dim STitle$(1 To 5), TDed$, DTitle$(1 To 5), AMAmt#
  Dim FTitle$
  Dim LastDed As Integer, DCnt As Integer
  Dim ETitle$, TErn$, LastErn As Integer, SumHeader2$
  Dim PayRegisterRptName$, RHandle As Integer
  Dim DHandle As Integer, DistributionRptName$
  Dim DistribuAcctRptName$
  Dim AHandle As Integer
  Dim DistAcctTotRptName$
  Dim ATHandle As Integer
  Dim DedDescRptName$
  Dim DDHandle As Integer
  Dim DistribuFundRptName$
  Dim FHandle As Integer
  Dim DistFundTotRptName$
  Dim TFHandle As Integer
  Dim DistFundEmpTotals$
  Dim FETHandle As Integer
  Dim NHandle As Integer, THandle As Integer
  Dim Cnt2 As Integer, cnt As Integer
  Dim Acct$, ENumOfAct As Integer, First As Integer
  Dim Start As Integer, Last As Integer
  Dim TotalSocAmt#, LastActive As Integer
  Dim Cnt3 As Integer, DistDif#, TotalMedAmt#
  Dim TotalRetAmt#, DLincCnt As Integer
  Dim Passed#(1 To 6), CCnt As Integer
  Dim FTax#, STax#, SSTax#, MTax#, RETTOT#, TNetPay#
  Dim RptTitle$, RegHrs#, VACHRS#, SICKHRS#, HOLHRS#
  Dim COMPHRS#, PerHRS#, TotalHrs#, TotEIC#, TotHrs#
  Dim TOTPaid#, TOTComp#, TRegWage#, TOTWage#, GPay#
  Dim SumDed$(1 To 5), SysHandle As Integer, DedAmt#
  Dim SumErn$, FF$, ARAmt#, CCnt3 As Integer
  Dim RegWage#, OTWage#, AddEarn#, DGPay#, ASAmt#
  Dim SplDedCnt As Integer, TempLine$(1 To 50), SFund$, zCnt As Integer
  Dim EmpSplCnt As Integer, ESCnt As Integer, SPctPCT#
  Dim ESplAmt#, GotIt As Boolean, CCnt01 As Integer, SDedAmt#
  Dim tFedTax#, tStaTax#, tMedTax#, aa As Integer
  Dim tSocTax#, tRetTax#, tEICPay#
  Dim Nextx As Integer, AddLine%
  Dim tripCnt As Integer, EIC As Double
  Dim NumOfDeds As Integer, dlm$, AcctCnt As Integer
  Dim DedDescTrimmed(1 To 50) As String * 8
  Dim TempMedTaxAmt#
  
  Dim CSocWHAcct$
  Dim CMedWHAcct$
  Dim CSocMatchAcct$
  Dim CMedMatchAcct$
  Dim CRetMatchAcct$
  Dim CFedWHAcct$
  Dim CStaWHAcct$
  Dim CRetWHAcct$
  Dim CDedAcct$
  Dim CPRNetAcct$
  Dim DWagesAcct$
  Dim DSocMatchAcct$
  Dim DMedMatchAcct$
  Dim DRetMatchAcct$
  Dim FundNumOnly$
  Dim FundAndAcctOnly$
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim TVCnt As Double
'  Dim FundLen As Integer
  Dim AcctLen As Integer
  Dim DetLen As Integer
  Dim ThisPR As Double
  Dim NextSpl As Integer
  Dim ActiveCnt As Integer
  
  On Error GoTo ErrorHandler
  
  ActiveCnt = 0
  dlm$ = "~"
  FF$ = Chr$(12)
  OpenPPDefaultFile PPDefaultHandle
  Get PPDefaultHandle, 1, PDR(1)
  Close PPDefaultHandle
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedCodeFileHandle
  ReDim DedCodeNums(1 To 1) As String '6/15/04
  For x = 1 To 50
    Get DedCodeFileHandle, x, DedRec
    DedCodes(x) = DedRec
    RSet DedDescTrimmed(x) = QPTrim$(DedRec.DCDESC1)
    If Len(QPTrim(DedRec.DCDESC1)) > 0 Then
      NumOfDeds = NumOfDeds + 1
      ReDim Preserve DedCodeNums(1 To NumOfDeds) As String '6/15/04
      DedCodeNums(NumOfDeds) = QPTrim$(DedRec.DCACCT1)
    End If
  Next x
  Close DedCodeFileHandle
  
  OpenErnCodeFile ErnCodeFileHandle
  For x = 1 To 3
    Get ErnCodeFileHandle, x, ErnCodes(x)
  Next x
  Close ErnCodeFileHandle
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec(1)
  Close SysHandle
  FACnt = SysRec(1).AcctCnt
  
'  Call GetAcctStruct(SysRec(1).CITIDIR, FundLen, AcctLen, DetLen)
  Call GetAcctStruct(CurrCitiPath, FundLen, AcctLen, DetLen)
  
  Image0$ = "###0"
  Image$ = "###0.00"
  Image3$ = "###,##0.00"
  Image4$ = "###.####"
  Image5$ = "###,##0.00"
  
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0
  GTaxFrn# = 0
  TotDebit# = 0
  TotCredit# = 0
  
  ArraySize = 1
  
  EmpActLen = Len(EmpAct(1))
  DistbSumAcctsLen = Len(DistbSumAccts(1))
  
  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))
  
  OpenEmpData1File Emp1Handle
  NumOfRecs = LOF(Emp1Handle) / Len(EmpRec1(1))
  Close Emp1Handle
  
  SalCnt = 0
  HrlCnt = 0
  
  NumOfWageAccts = 0
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  For x = 1 To NumOfRecs
    Get EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  PayRegisterRptName$ = "PRRPTS\REGISTERG.RPT"
  RHandle = FreeFile
  Open PayRegisterRptName$ For Output As RHandle
  
  DistributionRptName$ = "PRRPTS\DISTRIBU.RPT"
  DHandle = FreeFile
  Open DistributionRptName$ For Output As DHandle
  
  DistribuAcctRptName$ = "PRRPTS\DISTRIBUACCTNUMG.RPT"
  AHandle = FreeFile
  Open DistribuAcctRptName$ For Output As AHandle
  
  DistribuFundRptName$ = "PRRPTS\DISTRIBUFUNDNUM.RPT"
  FHandle = FreeFile
  Open DistribuFundRptName$ For Output As FHandle
  
  DistFundEmpTotals$ = "PRRPTS\DISTFUNDEMPTOTS.RPT"
  FETHandle = FreeFile
  Open DistFundEmpTotals$ For Output As FETHandle
  
  DistAcctTotRptName$ = "PRRPTS\DISTACCTTOT.RPT"
  ATHandle = FreeFile
  Open DistAcctTotRptName$ For Output As ATHandle
  
  DistFundTotRptName$ = "PRRPTS\DISTFUNDNUMTOTALS.RPT"
  TFHandle = FreeFile
  Open DistFundTotRptName$ For Output As TFHandle
  
  DedDescRptName$ = "PRRPTS\DEDDESC.RPT"
  DDHandle = FreeFile
  Open DedDescRptName$ For Output As DDHandle
  Nextx = 1
  For x = 1 To 5
    '                      0 to 49
    Print #DDHandle, DedDescTrimmed(Nextx); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 1); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 2); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 3); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 4); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 5); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 6); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 7); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 8); dlm;
    Print #DDHandle, DedDescTrimmed(Nextx + 9); dlm; x
    If x <> 5 Then
      If QPTrim$(DedDescTrimmed(Nextx + 10)) = "" Then Exit For
    End If
    Nextx = Nextx + 10
  Next x
  Close DDHandle
  
  OpenEmpData1File NHandle
  OpenTransWorkFile THandle
  
  FrmShowPctComp.Label1 = "Printing Earnings Distribution Register"
  FrmShowPctComp.Show ' , Me
  
  KillFile TempVoidFileName
  OpenTempVoidFile TVHandle
  
  For cnt = 1 To NumOfRecs
    Get THandle, IdxBuff(cnt), TransRec(1)
    If TransRec(1).TActive = True Then
      Get NHandle, IdxBuff(cnt), EmpRec1(1)
      GoSub SumAndPrintTime
      GoSub ParseDistributions
      ActiveCnt = ActiveCnt + 1
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
BadEmp:
  Next
  
  If ActiveCnt = 0 Then
    MsgBox "No employees have been designated for payroll processing."
    Close
    GoTo AltExit
  End If
  
  GoSub PrintDistTotal
  GoSub PrintSplitTotal
  
  Close THandle
  Close NHandle
  Close RHandle
  Close DHandle
  Close AHandle
  Close ATHandle
  Close FHandle
  Close FETHandle
  Close TFHandle
  Close TVHandle
  
  
  
'  Dim NumAccts As Integer
'
'  OpenTempVoidFile TVHandle
'  NumAccts = LOF(TVHandle) / Len(TempVoid)
'  For X = 1 To NumAccts
'    Get TVHandle, X, TempVoid
'      Debug.Print TempVoid.EmpNum
'      Debug.Print TempVoid.PPEGL + " PPE              " + CStr(TempVoid.PPEAmt)
'      Debug.Print TempVoid.PPETotGL + " PPE Total   " + CStr(TempVoid.PPETotAmt)
'      Debug.Print TempVoid.PRNetGL + " PRNET            " + CStr(TempVoid.PRNet)
'      Debug.Print TempVoid.SOCWHGL + " SOC Withholdings " + CStr(TempVoid.SOCWHAmt)
'      Debug.Print TempVoid.MEDWHGL + " MED Withholdings " + CStr(TempVoid.MEDWHAmt)
'      Debug.Print TempVoid.SOCMATCRGL + " SOC Match Liab   " + CStr(TempVoid.SOCMATCRAmt)
'      Debug.Print TempVoid.MEDMATCRGL + " MED Match Liab   " + CStr(TempVoid.MEDMATCRAmt)
'      Debug.Print TempVoid.FEDWHGL + " FED Withholdings " + CStr(TempVoid.FEDWHAmt)
'      Debug.Print TempVoid.STAWHGL + " STA Withholdings " + CStr(TempVoid.STAWHAmt)
'      Debug.Print TempVoid.RETWHGL + " RET Withholdings " + CStr(TempVoid.RETWHAmt)
'      Debug.Print TempVoid.RETMATCRGL + " RET Match Liab   " + CStr(TempVoid.RETMATCRAmt)
'      For CCnt3 = 1 To 50
'        If TempVoid.DedData(CCnt3).DAmt > 0 Then
'          Debug.Print TempVoid.DedData(CCnt3).DedGLNum + " Deduction        " + CStr(TempVoid.DedData(CCnt3).DAmt)
'        End If
'      Next CCnt3
'      Debug.Print TempVoid.WagesGL + "  Wages           " + CStr(TempVoid.WagesAmt)
'      Debug.Print TempVoid.SOCMATDBGL + " SOC Match        " + CStr(TempVoid.SOCMATDBAmt)
'      Debug.Print TempVoid.MEDMATDBGL + " MED Match        " + CStr(TempVoid.MEDMATDBAmt)
'      Debug.Print TempVoid.RETMATDBGL + " RET Match        " + CStr(TempVoid.RETMATDBAmt)
'  Next X
'  Close TVHandle
  
  'HERE:  register report files have been written to disk
  'GLIF START
  'if there is a GL transfer directory then make GLIF file
'  If Len(QPTrim$(SysRec(1).CITIDIR)) > 0 Then
'  MsgBox "Length of CurrCitiPath is " + CStr(Len(CurrCitiPath))
  If Len(CurrCitiPath) > 0 Then
    Passed#(1) = FTax#
    Passed#(2) = STax#
    Passed#(3) = SSTax#
    Passed#(4) = MTax#
    Passed#(5) = RETTOT#
    Passed#(6) = TNetPay#
    MakeGLIFFileSG SplitDeds(), DistbSumAccts(), TotEIC#, TotDeds#(), Passed#(), PrintGLRpt
    'unrem
  End If
  
'  Unload frmLoadingRpt
  If PrintGLRpt Then
    If GLCreditTotal <> GLDebitTotal Then
      frmBackGround.Show
      frmMessage.Label1.Caption = "The General Ledger Interface is OUT OF BALANCE."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      Unload frmBackGround
      MainLog "User warned that the GL Interface is OUT OF BALANCE (Debit Total = " + QPTrim$(Using$("$#,###,##0.00", GLDebitTotal)) + " and the Credit Total is " + QPTrim$(Using$("$#,###,##0.00", GLCreditTotal)) + ")."
    End If
    If GLError < 1 Then
      frmBackGround.Show
      frmMessage.Label1.Caption = "General Ledger number errors have been found in the GL Interface Report."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      Unload frmBackGround
      MainLog "User warned that the GL Interface Report has GL number errors."
    End If
    
    RptTitle$ = " Report"
    arGLRegister.Show
    frmLoadingRpt.Show
  End If
  
  RptTitle$ = "Payroll Distribution Report"
  arDistRegAcctNum.Show
  frmLoadingRpt.Show
  If RetCode = -1 Then GoTo AltExit
  
  RptTitle$ = "Payroll Register Report"
  arPRRegister.Show
  frmLoadingRpt.Show
  
  If RetCode = -1 Then GoTo AltExit
  
  
AltExit:
  
Exit Sub
  
SumAndPrintTime:
  RegHrs# = OldRound#(RegHrs# + TransRec(1).RegHrsWork)
  VACHRS# = OldRound#(VACHRS# + TransRec(1).VacUsed)
  SICKHRS# = OldRound#(SICKHRS# + TransRec(1).SickUsed)
  HOLHRS# = OldRound#(HOLHRS# + TransRec(1).HOLHOURS)
  COMPHRS# = OldRound#(COMPHRS# + TransRec(1).CompUsed)
  PerHRS# = OldRound#(PerHRS# + TransRec(1).PerHours)
  
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).PerHours)

  'added for EIC   6/07/94
  TotEIC# = OldRound#(TotEIC# + TransRec(1).EICAmt)
  
  '-=-=-=-=-=-=-=
  TotHrs# = OldRound#(TotHrs# + TransRec(1).OTHours)
  If TransRec(1).OTHrsPaid > 0 Then
    TOTPaid# = OldRound#(TOTPaid# + TransRec(1).OTHrsPaid)
  End If
  TOTComp# = OldRound#(TOTComp# + TransRec(1).OT2Comp)
  
  TRegWage# = OldRound#(TRegWage# + TransRec(1).TotRegWage)
  
  If TransRec(1).TotOTWage > 0 Then
    TOTWage# = OldRound#(TOTWage# + TransRec(1).TotOTWage)
  End If
  
  
  GPay# = OldRound#(GPay# + TransRec(1).GrossPay)
  SSTax# = OldRound#(SSTax# + TransRec(1).SocTaxAmt)
  MTax# = OldRound#(MTax# + TransRec(1).MedTaxAmt)
  FTax# = OldRound#(FTax# + TransRec(1).FedTaxAmt)
  STax# = OldRound#(STax# + TransRec(1).StaTaxAmt)
  If TransRec(1).RetireAmt > 0 Then
    RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
  End If
  
  TNetPay# = OldRound#(TNetPay# + TransRec(1).NetPay)
  GFedGross# = OldRound#(GFedGross# + TransRec(1).FedGrossPay)
  GStaGross# = OldRound#(GStaGross# + TransRec(1).StaGrossPay)
  GSocGross# = OldRound#(GSocGross# + TransRec(1).SocGrossPay)
  GRetGross# = OldRound#(GRetGross# + TransRec(1).RetGrossPay)
  GMedGross# = OldRound#(GMedGross# + TransRec(1).MedGrossPay)
  
  GTaxFrn# = OldRound#(GTaxFrn# + TransRec(1).TaxFring)
  
  LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
  LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)
  RSet BRat(1) = Using(Image3$, TransRec(1).BaseRate)
  RSet ORat(1) = Using(Image3$, TransRec(1).OTRate)
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsWork)
  RSet VHrs(1) = Using(Image$, TransRec(1).VacUsed)
  RSet SHrs(1) = Using(Image$, TransRec(1).SickUsed)
  RSet HHrs(1) = Using(Image$, TransRec(1).HOLHOURS)
  RSet CHrs(1) = Using(Image$, TransRec(1).CompUsed)
  RSet THrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  RSet TaxFrn(1) = Using(Image$, TransRec(1).TaxFring)
  RSet OTHrs(1) = Using(Image$, TransRec(1).OTHours)
  RSet PHrs(1) = Using(Image$, TransRec(1).PerHours)

  RSet OTPaid(1) = Using(Image$, TransRec(1).OTHrsPaid)
  RSet OTComp(1) = Using(Image$, TransRec(1).OT2Comp)
  
  'added for EIC     6/07/94
  RSet EEicP(1) = Using(Image3$, TransRec(1).EICAmt)
  
  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  '=======
  RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
  RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
  
  RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
  RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
  RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
  RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
  RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
  RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
  RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  
  For Cnt2 = 1 To NumOfDeds
    TotDeds#(Cnt2) = OldRound#(TotDeds#(Cnt2) + TransRec(1).DAmt(Cnt2))
  Next
  
  For Cnt2 = 1 To 3
    TotErns(Cnt2) = OldRound#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
NoMoreErns:
  Next
  '-------------------------------------------------------
  '                        0                               1                        2
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm; ENumb(1); dlm;
  '                  3
  Print #RHandle, EName(1); dlm;
  '
  If Len(QPTrim$(ErnCodes(1).ERNCODE1)) > 0 Then
  '                          4                         5
    Print #RHandle, Using$("###,###,##0.00", TransRec(1).EAmt(1)); dlm; QPTrim$(ErnCodes(1).ERNCODE1); dlm;
  Else
  '                 4        5
    Print #RHandle, ""; dlm; ""; dlm;
  End If
  
  If Len(QPTrim$(ErnCodes(2).ERNCODE1)) > 0 Then
  '                         6                              7
    Print #RHandle, Using$("###,###,##0.00", TransRec(1).EAmt(2)); dlm; QPTrim$(ErnCodes(2).ERNCODE1); dlm;
  Else
  '                  6        7
    Print #RHandle, ""; dlm; ""; dlm;
  End If
  
  If Len(QPTrim$(ErnCodes(3).ERNCODE1)) > 0 Then
  '                           8                              9
    Print #RHandle, Using$("###,###,##0.00", TransRec(1).EAmt(3)); dlm; QPTrim$(ErnCodes(3).ERNCODE1); dlm;
  Else
  '                 8        9
    Print #RHandle, ""; dlm; ""; dlm;
  End If
   
  '10 - 59
  For x = 1 To 50
    Print #RHandle, QPTrim$(DedCodes(x).DCDESC1); dlm;
  Next x
  '                 60            61            62             63
  Print #RHandle, BRat(1); dlm; ORat(1); dlm; TaxFrn(1); dlm; RHrs(1); dlm;
  '                 64            65            66            67
  Print #RHandle, VHrs(1); dlm; SHrs(1); dlm; HHrs(1); dlm; CHrs(1); dlm;
  '                 68            69             70              71
  Print #RHandle, PHrs(1); dlm; THrs(1); dlm; OTPaid(1); dlm; OTComp(1); dlm;
  '                 72           73              74             75
  Print #RHandle, EEicP(1); dlm; RErnP(1); dlm; OErnP(1); dlm; TotErns(1); dlm;
  '                   76               77              78             79
  Print #RHandle, TotErns(2); dlm; TotErns(3); dlm; GPayP(1); dlm; SSTaxP(1); dlm;
  '                  80            81             82              83
  Print #RHandle, MTaxP(1); dlm; FTaxP(1); dlm; STaxP(1); dlm; RetirP(1); dlm;
  '                  84
  Print #RHandle, NetPayP(1); dlm;
'  ' 85 - 134
  For x = 1 To 50
    If x > NumOfDeds Then
      Print #RHandle, ""; dlm;
    Else
      Print #RHandle, Using("###,###,##0.00", TransRec(1).DAmt(x)); dlm;
    End If
  Next x
  '                 135           136              137                138
  Print #RHandle, SalCnt; dlm; HrlCnt; dlm; GFedGross#; dlm; GStaGross#; dlm;
  '                   139               140            141
  Print #RHandle, GSocGross#; dlm; GRetGross#; dlm; GMedGross#; dlm;
  '142 -190
  For x = 1 To 49
    If x > NumOfDeds Then
      Print #RHandle, ""; dlm;
    Else
      Print #RHandle, Using("###,###,##0.00", TotDeds#(x)); dlm;
    End If
  Next x
  
  If 50 > NumOfDeds Then
  '                    191
      Print #RHandle, ""
  Else
  '                                 191
      Print #RHandle, Using("###,###,##0.00", TotDeds#(50))
  End If
Return
  
  '-----------------------------------------------------------------------
ParseDistributions:

  ReDim EmpAct(1 To 14) As DistWageRptType
  ENumOfAct = 0
  '                         50                               51
  Print #DHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
  
  '                  52             53            54            55
  Print #DHandle, ENumb(1); dlm; EName(1); dlm; BRat(1); dlm; ORat(1)
  
  'process wage distributions
  For Cnt2 = 1 To 8
    Acct$ = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    If Len(Acct$) > 0 And TransRec(1).TDist(Cnt2).DPct > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).Pct = TransRec(1).TDist(Cnt2).DPct
      EmpAct(ENumOfAct).RHrs = TransRec(1).TDist(Cnt2).DRHrs
      EmpAct(ENumOfAct).OHrs = TransRec(1).TDist(Cnt2).DOHrs
      EmpAct(ENumOfAct).RWage = TransRec(1).TDist(Cnt2).DRWage
      EmpAct(ENumOfAct).OWage = TransRec(1).TDist(Cnt2).DOWage
      EmpAct(ENumOfAct).GrossPay = OldRound#(EmpAct(ENumOfAct).RWage + EmpAct(ENumOfAct).OWage)
    End If
  Next
  'process earnings distributions 'start here 8/20/07
  For Cnt2 = 1 To 6
    Acct$ = QPTrim$(TransRec(1).EDist(Cnt2).EAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).AddEarn = TransRec(1).EDist(Cnt2).EAmt
      EmpAct(ENumOfAct).GrossPay = TransRec(1).EDist(Cnt2).EAmt
    End If
  Next
  
  'HERE: got all accts for this employee
  
  'sum accts
  
  First = 1
  Start = 1
  
  Last = ENumOfAct
 'Start here on 10/20/04
'purge and sum employee's dupe account distributions
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If EmpAct(First).Acct = EmpAct(Cnt2).Acct Then
        LSet EmpAct(Cnt2).Acct = ""             'blank acct num as a flag
        EmpAct(First).Pct = OldRound#(EmpAct(First).Pct + EmpAct(Cnt2).Pct)
        EmpAct(First).RHrs = OldRound#(EmpAct(First).RHrs + EmpAct(Cnt2).RHrs)
        EmpAct(First).OHrs = OldRound#(EmpAct(First).OHrs + EmpAct(Cnt2).OHrs)
        EmpAct(First).RWage = OldRound#(EmpAct(First).RWage + EmpAct(Cnt2).RWage)
        EmpAct(First).OWage = OldRound#(EmpAct(First).OWage + EmpAct(Cnt2).OWage)
        EmpAct(First).AddEarn = OldRound#(EmpAct(First).AddEarn + EmpAct(Cnt2).AddEarn)
        EmpAct(First).GrossPay = OldRound#(EmpAct(First).GrossPay + EmpAct(Cnt2).GrossPay)
      End If
    Next
Again:
    First = First + 1
  Loop Until First >= Last      '
  
  'calc percentages of matching amts to each account distribution

  For Cnt2 = 1 To ENumOfAct
    EmpAct(Cnt2).MATSocAmt = OldRound#(TransRec(1).MatchSocAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATMedAmt = OldRound#(TransRec(1).MatchMedAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATRetAmt = OldRound#(TransRec(1).MatchRetAmt * (EmpAct(Cnt2).Pct * 0.01))
  Next
  '---------------------------------------------------------------------------
'calc and adjust matching distribution amts, find last active account
  'adjust Social Amt
  Do
    TotalSocAmt# = 0
    For Cnt2 = 1 To 14
      TotalSocAmt# = OldRound#(TotalSocAmt# + EmpAct(Cnt2).MATSocAmt)
    Next
    If TotalSocAmt# = 0 Then GoTo SkipSocDist
    If TotalSocAmt# <> TransRec(1).MatchSocAmt Then
      For Cnt3 = 14 To 1 Step -1
        If EmpAct(Cnt3).MATSocAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalSocAmt# > TransRec(1).MatchSocAmt Then
        DistDif# = OldRound#(TotalSocAmt# - TransRec(1).MatchSocAmt)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt - DistDif#)
      ElseIf TotalSocAmt# < TransRec(1).MatchSocAmt Then
        DistDif# = OldRound#(TransRec(1).MatchSocAmt - TotalSocAmt#)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt + DistDif#)
      End If
    End If
  Loop Until TotalSocAmt# = TransRec(1).MatchSocAmt

SkipSocDist:
  'adjust Medicare Amt
  Do
    TotalMedAmt# = 0
    For Cnt2 = 1 To 8
      TotalMedAmt# = OldRound#(TotalMedAmt# + EmpAct(Cnt2).MATMedAmt)
    Next

    If TotalMedAmt# = 0 Then GoTo SkipMedDist
    If TotalMedAmt# <> OldRound#(TransRec(1).MatchMedAmt) Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATMedAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalMedAmt# > TransRec(1).MatchMedAmt Then
        DistDif# = OldRound#(TotalMedAmt# - TransRec(1).MatchMedAmt)
        EmpAct(LastActive).MATMedAmt = OldRound#(EmpAct(LastActive).MATMedAmt - DistDif#)
      ElseIf TotalMedAmt# < TransRec(1).MatchMedAmt Then
        DistDif# = OldRound#(TransRec(1).MatchMedAmt - TotalMedAmt#)
        EmpAct(LastActive).MATMedAmt = OldRound#(EmpAct(LastActive).MATMedAmt + DistDif#)
      End If
    End If
  Loop Until TotalMedAmt# = OldRound#(TransRec(1).MatchMedAmt)

SkipMedDist:
  'adjust Retire Amt
  Do
    TotalRetAmt# = 0
    For Cnt2 = 1 To 8
      TotalRetAmt# = OldRound#(TotalRetAmt# + EmpAct(Cnt2).MATRetAmt)
    Next
    If TotalRetAmt# = 0 Then GoTo SkipRetDist
    If TotalRetAmt# <> TransRec(1).MatchRetAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATRetAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalRetAmt# > TransRec(1).MatchRetAmt Then
        DistDif# = OldRound#(TotalRetAmt# - TransRec(1).MatchRetAmt)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt - DistDif#)
      ElseIf TotalRetAmt# < TransRec(1).MatchRetAmt Then
        DistDif# = OldRound#(TransRec(1).MatchRetAmt - TotalRetAmt#)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt + DistDif#)
      End If
    End If
  Loop Until TotalRetAmt# = TransRec(1).MatchRetAmt

SkipRetDist:
  
'------------------------------------------------------------------------------------
'print this employee's distributions
  AcctCnt = 0
  For Cnt2 = 1 To Last
 
    If Len(QPTrim$(EmpAct(Cnt2).Acct)) > 0 Then
      GoSub PrintEmpDist
    End If
  Next
  
  GoSub PrintEmpSubTotal
'look here

  AcctCnt = 0

  GoSub SplitDed
  
  ' sum to master distribution list
  For Cnt2 = 1 To Last          'process wage distributions
    Acct$ = QPTrim$(EmpAct(Cnt2).Acct)
    If Len(Acct$) > 0 Then
      If NumOfWageAccts > 0 Then
        For Cnt3 = 1 To NumOfWageAccts
          If Acct$ = QPTrim$(DistbSumAccts(Cnt3).Acct) Then
            DistbSumAccts(Cnt3).RWage = OldRound#(DistbSumAccts(Cnt3).RWage + EmpAct(Cnt2).RWage)
            DistbSumAccts(Cnt3).OWage = OldRound#(DistbSumAccts(Cnt3).OWage + EmpAct(Cnt2).OWage)
            DistbSumAccts(Cnt3).RHrs = OldRound#(DistbSumAccts(Cnt3).RHrs + EmpAct(Cnt2).RHrs)
            DistbSumAccts(Cnt3).OHrs = OldRound#(DistbSumAccts(Cnt3).OHrs + EmpAct(Cnt2).OHrs)
            DistbSumAccts(Cnt3).AddEarn = OldRound#(DistbSumAccts(Cnt3).AddEarn + EmpAct(Cnt2).AddEarn)
            DistbSumAccts(Cnt3).GrossPay = OldRound#(DistbSumAccts(Cnt3).GrossPay + EmpAct(Cnt2).GrossPay)
            DistbSumAccts(Cnt3).MATSocAmt = OldRound#(DistbSumAccts(Cnt3).MATSocAmt + EmpAct(Cnt2).MATSocAmt)
            DistbSumAccts(Cnt3).MATMedAmt = OldRound#(DistbSumAccts(Cnt3).MATMedAmt + EmpAct(Cnt2).MATMedAmt)
            DistbSumAccts(Cnt3).MATRetAmt = OldRound#(DistbSumAccts(Cnt3).MATRetAmt + EmpAct(Cnt2).MATRetAmt)
            Exit For
          End If
        Next
        If Cnt3 > NumOfWageAccts Then
          GoSub AddDistbSumAcct 'add new sum dist acct
        End If
      Else      'no previous sum accts. add new one
        GoSub AddDistbSumAcct   'add new sum dist acct
      End If
    End If
  Next
  
Return
  
AddDistbSumAcct:                'add amts to grand total acts summary
  
  NumOfWageAccts = NumOfWageAccts + 1
  If NumOfWageAccts > 1 Then
    ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  End If
  DistbSumAccts(NumOfWageAccts) = EmpAct(Cnt2)
  
Return
  
PrintEmpDist: 'arDistRegAcctNum
  LSet EDAct(1) = EmpAct(Cnt2).Acct
  
  AcctCnt = AcctCnt + 1
  RSet EDPct(1) = Using("##0.00", EmpAct(Cnt2).Pct)
  RSet EDRHrs(1) = Using(Image$, EmpAct(Cnt2).RHrs)
  RSet EDOHrs(1) = Using(Image$, EmpAct(Cnt2).OHrs)
  RSet EDRPay(1) = Using("##,##0.00", EmpAct(Cnt2).RWage)
  RSet EDOPay(1) = Using(Image$, EmpAct(Cnt2).OWage)
  RSet EDEarn(1) = Using(Image$, EmpAct(Cnt2).AddEarn)
  RSet EDGroP(1) = Using(Image$, EmpAct(Cnt2).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, EmpAct(Cnt2).MATSocAmt)
  RSet EDMAmt(1) = Using(Image$, EmpAct(Cnt2).MATMedAmt)
  RSet EDRAmt(1) = Using(Image$, EmpAct(Cnt2).MATRetAmt)
  
  '                         0                              1
  Print #AHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
  
  '                  2             3              4              5
  Print #AHandle, ENumb(1); dlm; EName(1); dlm; BRat(1); dlm; ORat(1); dlm;
  '               6                7                                 8                       9
  Print #AHandle, EDAct(1); dlm; EDPct(1); dlm; Using(Image$, EmpAct(Cnt2).RHrs); dlm; Using(Image$, EmpAct(Cnt2).OHrs); dlm;
  '                               10                                            11                                   12                                      13
  Print #AHandle, Using("##,##0.00", EmpAct(Cnt2).RWage); dlm; Using(Image$, EmpAct(Cnt2).OWage); dlm; Using(Image$, EmpAct(Cnt2).AddEarn); dlm; Using(Image$, EmpAct(Cnt2).GrossPay); dlm;
  '                               14                                       15                                      16                                    17
  Print #AHandle, Using(Image$, EmpAct(Cnt2).MATSocAmt); dlm; Using(Image$, EmpAct(Cnt2).MATMedAmt); dlm; Using(Image$, EmpAct(Cnt2).MATRetAmt); dlm; Last; dlm;
  
  DLineCnt = DLineCnt + 2
  '               18       19
  Print #AHandle, ""; dlm; ""; dlm;
  
  For x = 1 To 50
  ' 20 - 69
    Print #AHandle, ""; dlm;
  Next x
  '               70       71
  Print #AHandle, ""; dlm; ""; dlm;
  '               72       73
  Print #AHandle, ""; dlm; ""; dlm;
  '               74       75       76       77
  Print #AHandle, ""; dlm; ""; dlm; ""; dlm; ""


Return
'
PrintEmpSubTotal:
  AcctCnt = 0
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  RSet OTHrs(1) = Using(Image$, TransRec(1).OTHours)
  RSet EDRPay(1) = Using(Image$, TransRec(1).TotRegWage)
  RSet EDOPay(1) = Using(Image$, TransRec(1).TotOTWage)
  RSet EDEarn(1) = Using(Image$, TransRec(1).TotAdditEarn)
  RSet EDGroP(1) = Using(Image$, TransRec(1).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, TransRec(1).MatchSocAmt)
  RSet EDMAmt(1) = Using(Image$, TransRec(1).MatchMedAmt)
  RSet EDRAmt(1) = Using(Image$, TransRec(1).MatchRetAmt)
  
Return
 
PrintDistTotal:
  
  RegHrs# = 0
  TotHrs# = 0
  RegWage# = 0
  OTWage# = 0
  AddEarn# = 0
  DGPay# = 0
  
  If NumOfWageAccts > 1 Then
    SortT DistbSumAccts(), NumOfWageAccts
  End If
  
  'sort distribution account numbers (page 5)
  'print the Summary of ALL Distrubtion Accounts
  'ON LOCAL ERROR RESUME NEXT
  
  For cnt = 1 To NumOfWageAccts
    LSet EDAct(1) = DistbSumAccts(cnt).Acct
    If GPay# <= 0 Then Return
    RSet EDPct(1) = Using(Image4$, (DistbSumAccts(cnt).GrossPay / GPay#) * 100)
    RSet RHrs(1) = Using(Image3$, DistbSumAccts(cnt).RHrs)
    RegHrs# = OldRound#(RegHrs# + DistbSumAccts(cnt).RHrs)
    RSet OTHrs(1) = Using(Image3$, DistbSumAccts(cnt).OHrs)
    TotHrs# = OldRound#(TotHrs# + DistbSumAccts(cnt).OHrs)
    RSet EDRPay(1) = Using(Image3$, DistbSumAccts(cnt).RWage)
    RegWage# = OldRound#(RegWage# + DistbSumAccts(cnt).RWage)
    RSet EDOPay(1) = Using(Image3$, DistbSumAccts(cnt).OWage)
    OTWage# = OldRound#(OTWage# + DistbSumAccts(cnt).OWage)
    RSet EDEarn(1) = Using(Image3$, DistbSumAccts(cnt).AddEarn)
    AddEarn# = OldRound#(AddEarn# + DistbSumAccts(cnt).AddEarn)
    RSet EDGroP(1) = Using(Image3$, DistbSumAccts(cnt).GrossPay)
    DGPay# = OldRound#(DGPay# + DistbSumAccts(cnt).GrossPay)
    
    RSet EDSAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATSocAmt)
    ASAmt# = OldRound#(ASAmt# + DistbSumAccts(cnt).MATSocAmt)
    
    RSet EDMAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATMedAmt)
    AMAmt# = OldRound#(AMAmt# + DistbSumAccts(cnt).MATMedAmt)
    
    RSet EDRAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATRetAmt)
    ARAmt# = OldRound#(ARAmt# + DistbSumAccts(cnt).MATRetAmt)
    
    
    '***********Report Totals by Account Number**********
    
    '                         0                                             1
    If Len(QPTrim$(DistbSumAccts(cnt).Acct)) > 0 Then
      Print #ATHandle, DistbSumAccts(cnt).Acct; dlm; Using(Image4$, (DistbSumAccts(cnt).GrossPay / GPay#) * 100); dlm;
      '                         2                                           3
      Print #ATHandle, Using(Image3$, DistbSumAccts(cnt).RHrs); dlm; Using(Image3$, DistbSumAccts(cnt).OHrs); dlm;
      '                         4                                           5
      Print #ATHandle, Using(Image3$, DistbSumAccts(cnt).RWage); dlm; Using(Image3$, DistbSumAccts(cnt).OWage); dlm;
      '                         6                                            7
      Print #ATHandle, Using(Image3$, DistbSumAccts(cnt).AddEarn); dlm; Using(Image3$, DistbSumAccts(cnt).GrossPay); dlm;
      '                         8
      Print #ATHandle, Using(Image3$, DistbSumAccts(cnt).MATSocAmt); dlm;
      '                         9                                             10
      Print #ATHandle, Using(Image3$, DistbSumAccts(cnt).MATMedAmt); dlm; Using(Image3$, DistbSumAccts(cnt).MATRetAmt)
    
    End If
  Next
  
  '*********Report Totals for Distribution*********
  
  
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet OTHrs(1) = Using(Image3$, TotHrs#)
  RSet EDRPay(1) = Using(Image3$, RegWage#)
  RSet EDOPay(1) = Using(Image3$, OTWage#)
  RSet EDEarn(1) = Using(Image3$, AddEarn#)
  RSet EDGroP(1) = Using(Image3$, DGPay#)
  
  RSet EDSAmt(1) = Using(Image3$, ASAmt#)
  RSet EDMAmt(1) = Using(Image3$, AMAmt#)
  RSet EDRAmt(1) = Using(Image3$, ARAmt#)
  
Return
  
PrintSplitTotal:

  If SplDedCnt = 0 Then SplDedCnt = 1

  ReDim Preserve SplitDeds(1 To SplDedCnt) As SplitDedsType

  If SplDedCnt > 1 Then
    SortE SplitDeds(), SplDedCnt
  End If
    
'look here
  For CCnt3 = 1 To SplDedCnt
    If Len(QPTrim$(SplitDeds(CCnt3).Acct)) = 0 Then GoTo SkipSplLine
    '5/1 This is the total by fund
    LSet SplDed2(1) = SplitDeds(CCnt3).Acct
    TempLine$(CCnt3) = SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).FEDTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).STATAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).MEDTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).SOCTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).RetTax)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    
    '*********Report Totals by Fund...(last page)**********
    '                       485
    Print #TFHandle, SplitDeds(CCnt3).Acct; dlm; Using("###0.00", SplitDeds(CCnt3).FEDTAX); dlm; Using("###0.00", SplitDeds(CCnt3).STATAX); dlm;
    '
    Print #TFHandle, Using("###0.00", SplitDeds(CCnt3).MEDTAX); dlm; Using("###0.00", SplitDeds(CCnt3).SOCTAX); dlm; Using("###0.00", SplitDeds(CCnt3).RetTax); dlm;
    '
    For DCnt = 1 To 50 'goes thru 708
    '
      If DCnt > NumOfDeds Then
        Print #TFHandle, ""; dlm;
      Else
        Print #TFHandle, Using("###0.00", SplitDeds(CCnt3).DedAmts(DCnt)); dlm;
      End If
      RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).DedAmts(DCnt))
    Next
    Print #TFHandle, Unit(1).UFEMPR; dlm; MakeRegDate(PDR(1).PEREND)
    
SkipSplLine:
  Next
Return
  
SplitDed:
  EmpSplCnt = 0
  If ENumOfAct = 0 Then ENumOfAct = 1
  ReDim EmpSpl(1 To ENumOfAct) As SplitDedsType
  For CCnt3 = 1 To ENumOfAct
    SFund$ = Left$(EmpAct(CCnt3).Acct, 2)
    If SFund$ = "  " Then '??????
    Else
      GoSub CheckEmpSpl 'divide the next fund up for this employee
      GoSub CheckFund 'check to see if there are any more funds
    End If
  Next
  
  GoSub FixSplitBal
  
  '----------------------------New Void Check Code--6/15/2004-----------
  
  ReDim ThisSplAcct(1 To EmpSplCnt) As String
  '8/10/04
  NextSpl = 1
  For x = 1 To EmpSplCnt
    ThisSplAcct(x) = EmpSpl(x).Acct
  Next x
  
  TempVoid.EmpNum = QPTrim$(EmpRec1(1).EmpNo)
  TempVoid.NumOfAccts = 0
  
  For CCnt3 = 1 To ENumOfAct
    If Len(QPTrim$(EmpAct(CCnt3).Acct)) <> 0 Then
      TempVoid.NumOfAccts = TempVoid.NumOfAccts + 1
    End If
  Next CCnt3
  
  For CCnt3 = 1 To ENumOfAct
    ThisPR = EmpAct(CCnt3).GrossPay
    If Len(QPTrim$(EmpAct(CCnt3).Acct)) = 0 Then GoTo NoAcctNum
    EmpAct(CCnt3).Acct = ReplaceString(EmpAct(CCnt3).Acct, "-", "")
    FundNumOnly$ = QPTrim(Mid(EmpAct(CCnt3).Acct, 1, FundLen))
    
    FundAndAcctOnly$ = QPTrim(Mid(EmpAct(CCnt3).Acct, 1, FACnt))

    TempVoid.SOCMATDBAmt = EmpAct(CCnt3).MATSocAmt
    DSocMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).SOCEXP)
    DSocMatchAcct$ = ReplaceString(DSocMatchAcct$, "-", "")
    DSocMatchAcct$ = AddDashesToGLNumber(DSocMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.SOCMATDBGL = DSocMatchAcct$

    TempVoid.MEDMATDBAmt = EmpAct(CCnt3).MATMedAmt
    DMedMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).MEDEXP)
    DMedMatchAcct$ = ReplaceString(DMedMatchAcct$, "-", "")
    DMedMatchAcct$ = AddDashesToGLNumber(DMedMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.MEDMATDBGL = DMedMatchAcct$

    TempVoid.RETMATDBAmt = EmpAct(CCnt3).MATRetAmt
    DRetMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).RETEXP)
    DRetMatchAcct$ = ReplaceString(DRetMatchAcct$, "-", "")
    DRetMatchAcct$ = AddDashesToGLNumber(DRetMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.RETMATDBGL = DRetMatchAcct$

    DWagesAcct$ = QPTrim$(EmpAct(CCnt3).Acct)
    DWagesAcct$ = ReplaceString(DWagesAcct$, "-", "")
    DWagesAcct$ = AddDashesToGLNumber(DWagesAcct$, FundLen, AcctLen, DetLen)
    TempVoid.WagesGL = DWagesAcct$
    TempVoid.WagesAmt = EmpAct(CCnt3).GrossPay

    CMedMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).MEDLIAB)
    CMedMatchAcct$ = ReplaceString(CMedMatchAcct$, "-", "")
    CMedMatchAcct$ = AddDashesToGLNumber(CMedMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.MEDMATCRGL = CMedMatchAcct$
    TempVoid.MEDMATCRAmt = EmpAct(CCnt3).MATMedAmt
    
    CSocMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).SOCLIAB)
    CSocMatchAcct$ = ReplaceString(CSocMatchAcct$, "-", "")
    CSocMatchAcct$ = AddDashesToGLNumber(CSocMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.SOCMATCRGL = CSocMatchAcct$
    TempVoid.SOCMATCRAmt = EmpAct(CCnt3).MATSocAmt
    
    CRetMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).RETLIAB)
    CRetMatchAcct$ = ReplaceString(CRetMatchAcct$, "-", "")
    CRetMatchAcct$ = AddDashesToGLNumber(CRetMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.RETMATCRGL = CRetMatchAcct$
    TempVoid.RETMATCRAmt = EmpAct(CCnt3).MATRetAmt
    
    For NextSpl = 1 To EmpSplCnt
      If QPTrim$(ThisSplAcct(NextSpl)) = Mid(EmpAct(CCnt3).Acct, 1, FundLen) Then
       ' 8/17/04
        ThisSplAcct(NextSpl) = "" 'once read it cannot be read again
        TempVoid.FEDWHAmt = EmpSpl(NextSpl).FEDTAX
        ThisPR = ThisPR - EmpSpl(NextSpl).FEDTAX
        CFedWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(1).Acct)
        CFedWHAcct$ = ReplaceString(CFedWHAcct$, "-", "")
        CFedWHAcct$ = AddDashesToGLNumber(CFedWHAcct$, FundLen, AcctLen, DetLen)
        TempVoid.FEDWHGL = CFedWHAcct$
    
        TempVoid.STAWHAmt = EmpSpl(NextSpl).STATAX
        ThisPR = ThisPR - EmpSpl(NextSpl).STATAX
        CStaWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(2).Acct)
        CStaWHAcct$ = ReplaceString(CStaWHAcct$, "-", "")
        CStaWHAcct$ = AddDashesToGLNumber(CStaWHAcct$, FundLen, AcctLen, DetLen)
        TempVoid.STAWHGL = CStaWHAcct$
    
        TempVoid.MEDWHAmt = EmpSpl(NextSpl).MEDTAX
        ThisPR = ThisPR - EmpSpl(NextSpl).MEDTAX
        CMedWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(4).Acct)
        CMedWHAcct$ = ReplaceString(CMedWHAcct$, "-", "")
        CMedWHAcct$ = AddDashesToGLNumber(CMedWHAcct$, FundLen, AcctLen, DetLen)
        TempVoid.MEDWHGL = CMedWHAcct$
    
        TempVoid.SOCWHAmt = EmpSpl(NextSpl).SOCTAX
        ThisPR = ThisPR - EmpSpl(NextSpl).SOCTAX
        CSocWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(3).Acct)
        CSocWHAcct$ = ReplaceString(CSocWHAcct$, "-", "")
        CSocWHAcct$ = AddDashesToGLNumber(CSocWHAcct$, FundLen, AcctLen, DetLen)
        TempVoid.SOCWHGL = CSocWHAcct$
    
        TempVoid.RETWHAmt = EmpSpl(NextSpl).RetTax
        ThisPR = ThisPR - EmpSpl(NextSpl).RetTax
        CRetWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(5).Acct)
        CRetWHAcct$ = ReplaceString(CRetWHAcct$, "-", "")
        CRetWHAcct$ = AddDashesToGLNumber(CRetWHAcct$, FundLen, AcctLen, DetLen)
        TempVoid.RETWHGL = CRetWHAcct$
    
        For x = 1 To 50
          If EmpSpl(NextSpl).DedAmts(x) > 0 Then
            CDedAcct$ = FundNumOnly$ + DedCodeNums(x)
            CDedAcct$ = ReplaceString(CDedAcct$, "-", "")
            CDedAcct$ = AddDashesToGLNumber(CDedAcct$, FundLen, AcctLen, DetLen)
            TempVoid.DedData(x).DAmt = EmpSpl(NextSpl).DedAmts(x)
            ThisPR = ThisPR - EmpSpl(NextSpl).DedAmts(x)
            TempVoid.DedData(x).DedGLNum = CDedAcct$
            TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedDescTrimmed(x))
          Else
            TempVoid.DedData(x).DAmt = 0
            TempVoid.DedData(x).DedGLNum = ""
          End If
        Next x
        Exit For
      End If
    Next NextSpl
    
    If NextSpl = EmpSplCnt + 1 Then
      TempVoid.FEDWHAmt = 0
      TempVoid.STAWHAmt = 0
      TempVoid.MEDWHAmt = 0
      TempVoid.SOCWHAmt = 0
      TempVoid.RETWHAmt = 0
      For x = 1 To 50
        TempVoid.DedData(x).DAmt = 0
      Next x
    End If
      
    TempVoid.VoidFlag = False
    TempVoid.CheckAmt = 0
    TempVoid.CheckDate = 0
    TempVoid.CheckNum = "0"
    TempVoid.TransRec = 0
    TempVoid.PRNet = ThisPR
    TempVoid.PRNetGL = FundNumOnly + QPTrim$(SysRec(1).CashAcct)
    TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
    TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    TempVoid.Type = QPTrim$(SysRec(1).USEIMP)
    TempVoid.Pad = ""
    If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
      TempVoid.PPEAmt = ThisPR
      TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + FundNumOnly
      TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
      TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
      TempVoid.PPETotAmt = ThisPR
      TempVoid.PPETotGL = QPTrim$(SysRec(1).IDRACCT)
      TempVoid.PPETotGL = ReplaceString(TempVoid.PPETotGL, "-", "")
      TempVoid.PPETotGL = AddDashesToGLNumber(TempVoid.PPETotGL, FundLen, AcctLen, DetLen)
    Else
      TempVoid.PPEAmt = 0
      TempVoid.PPEGL = ""
      TempVoid.PPETotAmt = 0
      TempVoid.PPETotGL = ""
    End If
    TVCnt = TVCnt + 1
    Put TVHandle, TVCnt, TempVoid
    
NoAcctNum:
  Next CCnt3
  
  '----------------^^^^--------New Void Check Code--6/15/2004-----------
  
  ReDim EmpDedSum(1 To 55) As Double '17) As Double changed 5/1
  For CCnt3 = 1 To EmpSplCnt
    '               0        1
    Print #AHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm;
    '                  2             3            4        5
    Print #AHandle, ENumb(1); dlm; EName(1); dlm; ""; dlm; ""; dlm;
    '               6        7        8        9                                                     9
    Print #AHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    '               10       11       12       13                                                                       12                                      13
    Print #AHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    '               14       15       16       17                                15                                      16                                    17
    Print #AHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    
    LSet SplDed(1) = EmpSpl(CCnt3).Acct
    
    '********** PRINTS EMPLOYEE'S INDIVIDUAL FUND DATA ***********
    '                  18                       19
    Print #AHandle, EmpSplCnt; dlm; QPTrim$(EmpSpl(CCnt3).Acct); dlm;
    
    
    For x = 1 To 50
    ' 20 - 69
      If x > NumOfDeds Then
        Print #AHandle, ""; dlm;
        EmpDedSum(x) = OldRound(EmpDedSum(x) + EmpSpl(CCnt3).DedAmts(x))
      Else
        Print #AHandle, Using("##0.00", EmpSpl(CCnt3).DedAmts(x)); dlm;
        EmpDedSum(x) = OldRound(EmpDedSum(x) + EmpSpl(CCnt3).DedAmts(x))
      End If
    Next x
    For DCnt = 1 To 50 'NumOfDeds
      RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).DedAmts(DCnt))
      TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    Next
EndLine1:
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).FEDTAX)
    EmpDedSum(51) = OldRound(EmpDedSum(NumOfDeds + 1) + EmpSpl(CCnt3).FEDTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).STATAX)
    EmpDedSum(52) = OldRound(EmpDedSum(NumOfDeds + 2) + EmpSpl(CCnt3).STATAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).MEDTAX)
    EmpDedSum(53) = OldRound(EmpDedSum(NumOfDeds + 3) + EmpSpl(CCnt3).MEDTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).SOCTAX)
    EmpDedSum(54) = OldRound(EmpDedSum(NumOfDeds + 4) + EmpSpl(CCnt3).SOCTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).RetTax)
    EmpDedSum(55) = OldRound(EmpDedSum(NumOfDeds + 5) + EmpSpl(CCnt3).RetTax)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    '                               70                                         71
    Print #AHandle, Using("##0.00", EmpSpl(CCnt3).FEDTAX); dlm; Using("##0.00", EmpSpl(CCnt3).STATAX); dlm;
    '                              72                                             73
    Print #AHandle, Using("##0.00", EmpSpl(CCnt3).MEDTAX); dlm; Using("##0.00", EmpSpl(CCnt3).SOCTAX); dlm;
    '                              74                                      75                           76                          77
    Print #AHandle, Using("##0.00", EmpSpl(CCnt3).RetTax); dlm; QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm; EmpSplCnt
    
  Next
  
  '**********FUND TOTALS FOR EACH EMPLOYEE  ************
'  Print #FETHandle, ENumb(1); dlm;
'  For x = 1 To 50
'  '   377
'    Print #FETHandle, Using("##0.00", EmpDedSum(x)); dlm;
'  Next x
'
'  For x = 51 To 55
'    Print #FETHandle, Using("##0.00", EmpDedSum(x)); dlm;
'  Next x
'  Print #FETHandle, ""
  
Return
  
CheckEmpSpl:

  If EmpSplCnt = 0 Then
    GoSub AddEmpSpl
  End If
  GotIt = False
  
'look here .......dividing up the deduction amounts into
'applicable fund percentages by employee
  For ESCnt = 1 To EmpSplCnt
    If Left$(EmpSpl(ESCnt).Acct, 2) = SFund$ Then
      GotIt = True
      SPctPCT# = OldRound(EmpAct(CCnt3).Pct * 0.01) '6/10/03 added OldRound
      If SPctPCT# > 0 Then
        For DCnt = 1 To 50   '10/9
'          If TransRec(1).DAmt(DCnt) = 1.13 Then Stop
          ESplAmt# = OldRound(TransRec(1).DAmt(DCnt) * SPctPCT#)
'          ESplAmt# = TransRec(1).DAmt(DCnt) * SPctPCT#
          EmpSpl(ESCnt).DedAmts(DCnt) = OldRound(EmpSpl(ESCnt).DedAmts(DCnt) + ESplAmt#)
'          If EmpSpl(ESCnt).DedAmts(DCnt) = 231.38 Then Stop
'          EmpSpl(ESCnt).DedAmts(DCnt) = EmpSpl(ESCnt).DedAmts(DCnt) + ESplAmt#
        Next
        'look here....EmpSpl(ESCnt).FEDTAX = total federal tax for all funds,,,EmpSpl(ESCnt).STATAX same for state, etc.
'        EmpSpl(ESCnt).FEDTAX = OldRound(EmpSpl(ESCnt).FEDTAX + (OldRound(TransRec(1).FedTaxAmt) * SPctPCT#))
        EmpSpl(ESCnt).FEDTAX = OldRound(EmpSpl(ESCnt).FEDTAX + OldRound(TransRec(1).FedTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'        EmpSpl(ESCnt).STATAX = OldRound(EmpSpl(ESCnt).STATAX + (TransRec(1).StaTaxAmt) * SPctPCT#)
        EmpSpl(ESCnt).STATAX = OldRound(EmpSpl(ESCnt).STATAX + OldRound(TransRec(1).StaTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'        EmpSpl(ESCnt).MEDTAX = OldRound(EmpSpl(ESCnt).MEDTAX + (TransRec(1).MedTaxAmt) * (EmpAct(CCnt3).Pct * 0.01))
        EmpSpl(ESCnt).MEDTAX = OldRound(EmpSpl(ESCnt).MEDTAX + OldRound(TransRec(1).MedTaxAmt * SPctPCT#)) '6/10/03
        
'        TempMedTaxAmt# = OldRound(TransRec(1).MedTaxAmt * (EmpAct(CCnt3).Pct * 0.01))
'        TempMedTaxAmt# = OldRound(TransRec(1).MedTaxAmt * EmpAct(CCnt3).Pct * 0.01) '6/10/03 changed from line above
'        EmpSpl(ESCnt).MEDTAX = OldRound(EmpSpl(ESCnt).MEDTAX + TempMedTaxAmt#)
        
'        EmpSpl(ESCnt).SOCTAX = OldRound(EmpSpl(ESCnt).SOCTAX + (TransRec(1).SocTaxAmt) * SPctPCT#)
        EmpSpl(ESCnt).SOCTAX = OldRound(EmpSpl(ESCnt).SOCTAX + OldRound(TransRec(1).SocTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'        EmpSpl(ESCnt).RetTax = OldRound(EmpSpl(ESCnt).RetTax + (TransRec(1).RetireAmt) * SPctPCT#)
        EmpSpl(ESCnt).RetTax = OldRound(EmpSpl(ESCnt).RetTax + OldRound(TransRec(1).RetireAmt * SPctPCT#)) '6/10/03 changed from line above
'        EmpSpl(ESCnt).EICPay = OldRound(EmpSpl(ESCnt).EICPay + (TransRec(1).EICAmt) * SPctPCT#)
        EmpSpl(ESCnt).EICPay = OldRound(EmpSpl(ESCnt).EICPay + OldRound(TransRec(1).EICAmt * SPctPCT#)) '6/10/03 changed from line above
      End If
    End If
  Next
  
  If Not GotIt Then 'got an SFund value that doesn't match any Left$(EmpSpl(ESCnt).Acct, 2)
    GoSub AddEmpSpl
    GoTo CheckEmpSpl
  Else
    GotIt = False 'reset for the next fund
  End If

Return
 
AddEmpSpl:
  EmpSplCnt = EmpSplCnt + 1
  EmpSpl(EmpSplCnt).Acct = SFund$
Return
  
CheckFund:
  If SplDedCnt = 0 Then 'first time thru
    GoSub AddNewFund 'counts funds and assigns fund numbers to
'   to SplitDeds(SplDedCnt).Acct = SFund$
  End If
  
  GotIt = False
  Nextx = 1
  tripCnt = 1
  For CCnt01 = 1 To SplDedCnt 'this SplDedCnt increases dynamically
    If Left$(SplitDeds(CCnt01).Acct, 2) = SFund$ Then
      GotIt = True
      SPctPCT# = OldRound(EmpAct(CCnt3).Pct * 0.01) '6/10/03 added OldRound
      LSet EDAct(1) = SFund$
'      TempLine$(Nextx) = ""
      
      For DCnt = 1 To NumOfDeds
        SDedAmt# = OldRound(TransRec(1).DAmt(DCnt) * SPctPCT#)
        SplitDeds(CCnt01).DedAmts(DCnt) = OldRound(SplitDeds(CCnt01).DedAmts(DCnt) + SDedAmt#)
      Next
      
'      SplitDeds(CCnt01).FEDTAX = OldRound(SplitDeds(CCnt01).FEDTAX + OldRound(TransRec(1).FedTaxAmt) * SPctPCT#)
      SplitDeds(CCnt01).FEDTAX = OldRound(SplitDeds(CCnt01).FEDTAX + OldRound(TransRec(1).FedTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'      SplitDeds(CCnt01).STATAX = OldRound(SplitDeds(CCnt01).STATAX + (TransRec(1).StaTaxAmt) * SPctPCT#)
      SplitDeds(CCnt01).STATAX = OldRound(SplitDeds(CCnt01).STATAX + OldRound(TransRec(1).StaTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'      SplitDeds(CCnt01).MEDTAX = OldRound(SplitDeds(CCnt01).MEDTAX + (TransRec(1).MedTaxAmt) * SPctPCT#)
      SplitDeds(CCnt01).MEDTAX = OldRound(SplitDeds(CCnt01).MEDTAX + OldRound(TransRec(1).MedTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'      SplitDeds(CCnt01).SOCTAX = OldRound(SplitDeds(CCnt01).SOCTAX + (TransRec(1).SocTaxAmt) * SPctPCT#)
      SplitDeds(CCnt01).SOCTAX = OldRound(SplitDeds(CCnt01).SOCTAX + OldRound(TransRec(1).SocTaxAmt * SPctPCT#)) '6/10/03 changed from line above
'      SplitDeds(CCnt01).RetTax = OldRound(SplitDeds(CCnt01).RetTax + (TransRec(1).RetireAmt) * SPctPCT#)
      SplitDeds(CCnt01).RetTax = OldRound(SplitDeds(CCnt01).RetTax + OldRound(TransRec(1).RetireAmt * SPctPCT#)) '6/10/03 changed from line above
'      SplitDeds(CCnt01).EICPay = OldRound(SplitDeds(CCnt01).EICPay + (TransRec(1).EICAmt) * SPctPCT#)
      SplitDeds(CCnt01).EICPay = OldRound(SplitDeds(CCnt01).EICPay + OldRound(TransRec(1).EICAmt * SPctPCT#)) '6/10/03 changed from line above
      Exit For
    End If
  Next

  If Not GotIt Then 'we've been thru the above loop with no match
  'so add the next fund
    GoSub AddNewFund
    GoTo CheckFund
  Else
    GotIt = False
  End If
  
Return

AddNewFund:
  SplDedCnt = SplDedCnt + 1
  If SplDedCnt > ArraySize Then
    ArraySize = ArraySize + 5
    ReDim Preserve SplitDeds(1 To ArraySize) As SplitDedsType
  End If
  
  SplitDeds(SplDedCnt).Acct = SFund$
Return
  
FixSplitBal:

  tFedTax# = 0
  tStaTax# = 0
  tMedTax# = 0
  tSocTax# = 0
  tRetTax# = 0
  tEICPay# = 0
  
  For CCnt = 1 To 50 'NumOfDeds
    DedAmt# = 0
    For aa = 1 To EmpSplCnt
      DedAmt# = OldRound(DedAmt# + EmpSpl(aa).DedAmts(CCnt))
'      If DedAmt > 0 Then Stop
'      If DedAmt = 1.12 Or DedAmt = 1.13 Then Stop
'      If DedAmt > 0 Then Stop
    Next
    If DedAmt# > TransRec(1).DAmt(CCnt) Then
      DistDif# = OldRound(DedAmt# - TransRec(1).DAmt(CCnt))
      EmpSpl(EmpSplCnt).DedAmts(CCnt) = OldRound(EmpSpl(EmpSplCnt).DedAmts(CCnt) - DistDif#)
      SplitDeds(EmpSplCnt).DedAmts(CCnt) = OldRound(SplitDeds(EmpSplCnt).DedAmts(CCnt) - DistDif#)
    ElseIf DedAmt# < TransRec(1).DAmt(CCnt) Then
      DistDif# = OldRound(TransRec(1).DAmt(CCnt) - DedAmt#)
      EmpSpl(EmpSplCnt).DedAmts(CCnt) = OldRound(EmpSpl(EmpSplCnt).DedAmts(CCnt) + DistDif#)
      SplitDeds(EmpSplCnt).DedAmts(CCnt) = OldRound(SplitDeds(EmpSplCnt).DedAmts(CCnt) + DistDif#)
    End If
  Next
  
  For aa = 1 To EmpSplCnt
    EmpSpl(aa).Acct = EmpSpl(aa).Acct
    tFedTax# = OldRound(tFedTax# + EmpSpl(aa).FEDTAX)
    tStaTax# = OldRound(tStaTax# + EmpSpl(aa).STATAX)
    tMedTax# = OldRound(tMedTax# + EmpSpl(aa).MEDTAX)
    tSocTax# = OldRound(tSocTax# + EmpSpl(aa).SOCTAX)
    tRetTax# = OldRound(tRetTax# + EmpSpl(aa).RetTax)
    tEICPay# = OldRound(tEICPay# + EmpSpl(aa).EICPay)
  Next
  
  If tFedTax# > TransRec(1).FedTaxAmt Then
    DistDif# = OldRound(tFedTax# - TransRec(1).FedTaxAmt)
    EmpSpl(EmpSplCnt).FEDTAX = OldRound(EmpSpl(EmpSplCnt).FEDTAX - DistDif#)
    SplitDeds(EmpSplCnt).FEDTAX = OldRound(SplitDeds(EmpSplCnt).FEDTAX - DistDif#)
  ElseIf tFedTax# < TransRec(1).FedTaxAmt Then
    DistDif# = OldRound(TransRec(1).FedTaxAmt - tFedTax#)
    EmpSpl(EmpSplCnt).FEDTAX = OldRound(EmpSpl(EmpSplCnt).FEDTAX + DistDif#)
    SplitDeds(EmpSplCnt).FEDTAX = OldRound#(SplitDeds(EmpSplCnt).FEDTAX + DistDif#)
  End If
  
  If tStaTax# > TransRec(1).StaTaxAmt Then
    DistDif# = OldRound(tStaTax# - TransRec(1).StaTaxAmt)
    EmpSpl(EmpSplCnt).STATAX = OldRound(EmpSpl(EmpSplCnt).STATAX - DistDif#)
    SplitDeds(EmpSplCnt).STATAX = OldRound(SplitDeds(EmpSplCnt).STATAX - DistDif#)
  ElseIf tStaTax# < TransRec(1).StaTaxAmt Then
    DistDif# = OldRound(TransRec(1).StaTaxAmt - tStaTax#)
    EmpSpl(EmpSplCnt).STATAX = OldRound(EmpSpl(EmpSplCnt).STATAX + DistDif#)
    SplitDeds(EmpSplCnt).STATAX = OldRound(SplitDeds(EmpSplCnt).STATAX + DistDif#)
  End If
  
  If tMedTax# > TransRec(1).MedTaxAmt Then
    DistDif# = OldRound(tMedTax# - TransRec(1).MedTaxAmt)
    EmpSpl(EmpSplCnt).MEDTAX = OldRound(EmpSpl(EmpSplCnt).MEDTAX - DistDif#)
    SplitDeds(EmpSplCnt).MEDTAX = OldRound(SplitDeds(EmpSplCnt).MEDTAX - DistDif#)
  ElseIf tMedTax# < TransRec(1).MedTaxAmt Then
    DistDif# = OldRound(TransRec(1).MedTaxAmt - tMedTax#)
    EmpSpl(EmpSplCnt).MEDTAX = OldRound(EmpSpl(EmpSplCnt).MEDTAX + DistDif#)
    SplitDeds(EmpSplCnt).MEDTAX = OldRound(SplitDeds(EmpSplCnt).MEDTAX + DistDif#)
  End If
  
  If tSocTax# > TransRec(1).SocTaxAmt Then
    DistDif# = OldRound(tSocTax# - TransRec(1).SocTaxAmt)
    EmpSpl(EmpSplCnt).SOCTAX = OldRound(EmpSpl(EmpSplCnt).SOCTAX - DistDif#)
    SplitDeds(EmpSplCnt).SOCTAX = OldRound(SplitDeds(EmpSplCnt).SOCTAX - DistDif#)
  ElseIf tSocTax# < TransRec(1).SocTaxAmt Then
    DistDif# = OldRound(TransRec(1).SocTaxAmt - tSocTax#)
    EmpSpl(EmpSplCnt).SOCTAX = OldRound(EmpSpl(EmpSplCnt).SOCTAX + DistDif#)
    SplitDeds(EmpSplCnt).SOCTAX = OldRound(SplitDeds(EmpSplCnt).SOCTAX + DistDif#)
  End If
  
  If tRetTax# > TransRec(1).RetireAmt Then
    DistDif# = OldRound(tRetTax# - TransRec(1).RetireAmt)
    EmpSpl(EmpSplCnt).RetTax = OldRound(EmpSpl(EmpSplCnt).RetTax - DistDif#)
    SplitDeds(EmpSplCnt).RetTax = OldRound(SplitDeds(EmpSplCnt).RetTax - DistDif#)
  ElseIf tRetTax# < TransRec(1).RetireAmt Then
    DistDif# = OldRound(TransRec(1).RetireAmt - tRetTax#)
    EmpSpl(EmpSplCnt).RetTax = OldRound(EmpSpl(EmpSplCnt).RetTax + DistDif#)
    SplitDeds(EmpSplCnt).RetTax = OldRound(SplitDeds(EmpSplCnt).RetTax + DistDif#)
  End If
  
  If tEICPay# > TransRec(1).EICAmt Then
    DistDif# = OldRound(tEICPay# - TransRec(1).EICAmt)
    EmpSpl(EmpSplCnt).EICPay = OldRound(EmpSpl(EmpSplCnt).EICPay - DistDif#)
    SplitDeds(EmpSplCnt).EICPay = OldRound(SplitDeds(EmpSplCnt).EICPay - DistDif#)
  ElseIf tEICPay# < TransRec(1).EICAmt Then
    DistDif# = OldRound(TransRec(1).EICAmt - tEICPay#)
    EmpSpl(EmpSplCnt).EICPay = OldRound(EmpSpl(EmpSplCnt).EICPay + DistDif#)
    SplitDeds(EmpSplCnt).EICPay = OldRound(SplitDeds(EmpSplCnt).EICPay + DistDif#)
  End If
  
Return
  
ErrorHandler:
  Close
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Sub MakeGLIFFileSG(SplitDeds() As SplitDedsType, DistbSumAccts() As DistWageRptType, TotEIC#, TotDeds#(), Passed#(), PrintGLRpt)

  Dim FEDWith$, MEDWith$, MEDMatLib$, SOCWith$
  Dim SOCMatLib$, RETWith$, RETMatLib$, STAWith$
  Dim MEDMat$, SOCMat$, RETMat$
  Dim Image3$, FACnt As Integer
  Dim NumOfWageAccts As Integer
  Dim SplDedCnt As Integer, GLIFTDate$
  Dim FedAcct$, StaAcct$, SocAcct$, MedAcct$
  Dim RetAcct$, GLIFSource$, cnt As Integer
  Dim GLIFArrayAdj As Integer, NextAcct As Integer
  Dim SysCash$, ZZCnt As Integer
  Dim TAcct$, NumDFunds As Integer
  Dim CurrFund$, FirstC As Integer
  Dim LastC As Integer, NumCFunds As Integer
  Dim TotalFunds As Integer, ThisFund As Integer
  Dim First As Integer, Start As Integer, Last As Integer
  Dim Cnt2 As Integer, FundCash$
  Dim TotalGLIFS As Integer, THandle As Integer
  Dim AcctNum$, SOCEXP$, MEDEXP$, RETEXP$
  Dim SOCLIAB$, MEDLIAB$, RETLIAB$
  Dim GLIFRecLen As Integer, RecNo&
  Dim GHandle As Integer, RptName$
  Dim TFHandle As Integer, TFRptName$
  Dim TotDebit#
  Dim GLIdxRecLen As Integer, TotCredit#
  Dim OutOfOrder As Boolean, AcctOk As Boolean
  Dim ActualAccts As Integer, NoAcctNum As Integer
  Dim ArraySize As Integer, GLIdxName$
  Dim Lines As Integer
  Dim Fund$, FDebit#, FCredit#, NFund$, FF$
  Dim GLHandle As Integer
  Dim GRHandle As Integer
  Dim y As Integer, z As Integer '8/10/04
  Dim Thisx As Integer
  Dim TOTEICFORRPT As Double
  
  FundCount = 0
  GLDebitTotal = 0
  GLCreditTotal = 0
  GLError = 0
  FEDWith$ = "FED Withholdings"      '    3
  MEDWith$ = "MED Withholdings"      '    4
  MEDMatLib$ = "MED Match Liab"      '    5
  SOCWith$ = "SOC Withholdings"      '    1
  SOCMatLib$ = "SOC Match Liab"      '    2
  RETWith$ = "RET Withholdings"      '    6
  RETMatLib$ = "RET Match Liab"      '    7
  STAWith$ = "STA Withholdings"      '    8

  MEDMat$ = "MED Match"
  SOCMat$ = "SOC Match"
  RETMat$ = "RET Match"              '

  ReDim SysRec(1) As RegDSysFileRecType
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType '8/10/04
  ReDim Unit(1) As UnitFileRecType
  ReDim EDesc(1) As String * 21
  ReDim Dash0(1) As String * 69
  LSet Dash0(1) = String$(69, "-")
  Dim FileHandle As Integer, Trans$
  Dim DedHandle As Integer, x As Integer
  Dim PDRHandle As Integer
  Dim SysHandle As Integer, Max As Integer
  Dim PPDefaultHandle As Integer
  Dim JFlag As Boolean, TotalAccts As Integer
  Dim GLIFRecName$, NumOfDeds As Integer
  Dim dlm$
  ReDim TotDedAmt(1 To 50) As Double '8/13/04
  Dim CIDebit#
  Dim CICredit#
  ReDim FundArray(1 To 1) As String '8/18/04
  Dim ThisDesc$ '8/18/04
  
  TOTFEDTAX = 0
'  TotEIC = 0 '5/6/09
  TOTMEDTAX = 0
  TOTSOCTAX = 0
  TOTSTATAX = 0
  TOTRetTax = 0
  TOTMEDMat = 0
  TOTSOCMat = 0
  TOTRETMat = 0
'  MsgBox "In MakeGLIFFileSG"
  dlm = "~"
  FF$ = Chr$(12)
  
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedHandle
  For x = 1 To 50
    Get DedHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedHandle
  
  OpenPPDefaultFile PDRHandle
  Get PDRHandle, 1, PDR(1)
  Close PDRHandle
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec(1)
  Close SysHandle
  
  Image3$ = "###,##0.00"

  FACnt = SysRec(1).AcctCnt 'Fund and Acct length

  NumOfWageAccts = UBound(DistbSumAccts)

  SplDedCnt = UBound(SplitDeds)

  'squeeze out all the "-" out of Acct numbers
  For cnt = 1 To NumOfWageAccts
    ReplaceString DistbSumAccts(cnt).Acct, "-", ""
  Next

  'change Period ending date to nicks format
  GLIFTDate$ = MakeRegDate(PDR(1).PEREND)

  ReplaceString GLIFTDate$, "-", "/"
  ReplaceString GLIFTDate$, "1994", "94"
  ReplaceString GLIFTDate$, "1995", "95"
  ReplaceString GLIFTDate$, "1996", "96"
  ReplaceString GLIFTDate$, "1997", "97"
  ReplaceString GLIFTDate$, "1998", "98"
  ReplaceString GLIFTDate$, "1999", "99"
  ReplaceString GLIFTDate$, "2000", "00"
  ReplaceString GLIFTDate$, "2001", "01"
  ReplaceString GLIFTDate$, "2002", "02"
  ReplaceString GLIFTDate$, "2003", "03"
  ReplaceString GLIFTDate$, "2004", "04"
  ReplaceString GLIFTDate$, "2005", "05"
  ReplaceString GLIFTDate$, "2006", "06"
  ReplaceString GLIFTDate$, "2007", "07"
  ReplaceString GLIFTDate$, "2008", "08"
  ReplaceString GLIFTDate$, "2009", "09"

  FedAcct$ = SysRec(1).Liab(1).Acct
  ReplaceString FedAcct$, "-", ""
  StaAcct$ = SysRec(1).Liab(2).Acct
  ReplaceString StaAcct$, "-", ""
  SocAcct$ = SysRec(1).Liab(3).Acct
  ReplaceString SocAcct$, "-", ""
  MedAcct$ = SysRec(1).Liab(4).Acct
  ReplaceString MedAcct$, "-", ""
  RetAcct$ = SysRec(1).Liab(5).Acct
  ReplaceString RetAcct$, "-", ""

  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$

  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  GLIFArrayAdj = 56
  GLIFArrayAdj = SplDedCnt * GLIFArrayAdj

  ArraySize = NumOfWageAccts + GLIFArrayAdj + SplDedCnt

  ReDim GLIFRec(1 To ArraySize) As GLIFDataType14 '8/31/07

  For cnt = 1 To NumOfWageAccts
    GLIFRec(cnt).TranAcct = DistbSumAccts(cnt).Acct
    GLIFRec(cnt).TranDate = GLIFTDate$
    GLIFRec(cnt).TranDesc = "Wages"
    GLIFRec(cnt).CrAmt = 0
    GLIFRec(cnt).DrAmt = DistbSumAccts(cnt).GrossPay
    GLIFRec(cnt).Source = GLIFSource$
    GLIFRec(cnt).FromFlag = "W"
  Next
  NextAcct = cnt

'06-02-95 split deductions
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + FedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = FEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).FEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt
'look here
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + StaAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = STAWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).STATAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + SocAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = SOCWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).SOCTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + MedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = MEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).MEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + RetAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = RETWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).RetTax
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  'jumpthis:

  NextAcct = NextAcct + cnt

  'good to here ;maybe

  ReplaceString SysRec(1).CashAcct, "-", ""

  SysCash$ = QPTrim$(SysRec(1).CashAcct)

  If TotEIC# > 0 Then
    ReDim EICGLIFRec(1 To (SplDedCnt * 2)) As GLIFDataType14
    For cnt = 1 To SplDedCnt
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt).Acct) + FedAcct$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = 0
      EICGLIFRec(cnt).DrAmt = SplitDeds(cnt).EICPay
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next

    For cnt = SplDedCnt + 1 To SplDedCnt * 2
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt - SplDedCnt).Acct) + SysCash$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = SplitDeds(cnt - SplDedCnt).EICPay
      EICGLIFRec(cnt).DrAmt = 0
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next

  End If
  For ZZCnt = 0 To (SplDedCnt - 1)
    For cnt = 1 To NumOfDeds
      ReplaceString DedCodes(cnt).DCACCT1, "-", ""
      TAcct$ = QPTrim$(SplitDeds(ZZCnt + 1).Acct)
      TAcct$ = TAcct$ + QPTrim$(DedCodes(cnt).DCACCT1)
      GLIFRec(NextAcct + (cnt - 1)).TranAcct = TAcct$
      GLIFRec(NextAcct + (cnt - 1)).TranDate = GLIFTDate$
      GLIFRec(NextAcct + (cnt - 1)).TranDesc = QPTrim$(DedCodes(cnt).DCDESC1) '"Deductions"
      GLIFRec(NextAcct + (cnt - 1)).CrAmt = SplitDeds(ZZCnt + 1).DedAmts(cnt)
      GLIFRec(NextAcct + (cnt - 1)).DrAmt = 0
      GLIFRec(NextAcct + (cnt - 1)).Source = GLIFSource$
      GLIFRec(NextAcct + (cnt - 1)).FromFlag = "D"
    Next
    NextAcct = NextAcct + cnt
  Next

  ReDim DFunds$(1 To NumOfWageAccts)

  NumDFunds = 1
  DFunds$(1) = Left$(DistbSumAccts(1).Acct, 2)
  For cnt = 1 To NumOfWageAccts - 1
    If Left$(DistbSumAccts(cnt).Acct, 2) <> Left$(DistbSumAccts(cnt + 1).Acct, 2) Then
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      DFunds$(NumDFunds) = Left$(DistbSumAccts(cnt + 1).Acct, 2)
    End If
  Next

  ReDim Preserve DFunds$(1 To NumDFunds)

  ReDim DFund(1 To NumDFunds) As FundType

  For cnt = 1 To NumOfWageAccts
    CurrFund$ = Left$(DistbSumAccts(cnt).Acct, 2)
    For ThisFund = 1 To NumDFunds
      If CurrFund$ = DFunds$(ThisFund) Then
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = OldRound(DFund(ThisFund).Debit + DistbSumAccts(cnt).GrossPay)

        Exit For
      End If
    Next
  Next

  ''all gross pay by funds here!!

  ReDim CFunds$(1 To GLIFArrayAdj)              '17)

  FirstC = NumOfWageAccts + 1
  LastC = NumOfWageAccts + GLIFArrayAdj + 1     '17

  NumCFunds = 1
  CFunds$(1) = Left$(GLIFRec(FirstC).TranAcct, 2)

  For cnt = FirstC To LastC - 1
    If Left$(GLIFRec(cnt).TranAcct, 2) <> Left$(GLIFRec(cnt + 1).TranAcct, 2) Then
      If Len(QPTrim$(GLIFRec(cnt + 1).TranAcct)) Then
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        CFunds$(NumCFunds) = Left$(GLIFRec(cnt + 1).TranAcct, 2)
      End If
    End If
  Next
    
  ReDim Preserve CFunds$(1 To NumCFunds)
  ReDim CFund(1 To NumCFunds) As FundType
  For cnt = FirstC To LastC - 1
    CurrFund$ = Left$(GLIFRec(cnt).TranAcct, 2)
    For ThisFund = 1 To NumCFunds
      If CurrFund$ = CFunds$(ThisFund) Then
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = OldRound(CFund(ThisFund).Credit + GLIFRec(cnt).CrAmt)
        Exit For
      End If
    Next
  Next

  ''combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  ReDim AllFunds(1 To TotalFunds) As FundType
  ThisFund = 1
  For cnt = 1 To NumDFunds
    AllFunds(ThisFund) = DFund(cnt)
    ThisFund = ThisFund + 1
  Next

  For cnt = 1 To NumCFunds
    AllFunds(ThisFund) = CFund(cnt)
    ThisFund = ThisFund + 1
  Next

  '
  SortF AllFunds(), TotalFunds
  'combine Debits and Credits for same fund numbers

  First = 1
  Start = 1
  Last = TotalFunds
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If AllFunds(First).FundCode = AllFunds(Cnt2).FundCode Then
        LSet AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = OldRound#(AllFunds(First).Credit + AllFunds(Cnt2).Credit)

        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = OldRound#(AllFunds(First).Debit + AllFunds(Cnt2).Debit)

        AllFunds(Cnt2).Debit = 0
      End If
    Next
    First = First + 1
  Loop Until First >= Last      '

  'calc net difference for funds
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(AllFunds(cnt).FundCode)) Then
      AllFunds(cnt).Net = OldRound#(AllFunds(cnt).Debit - AllFunds(cnt).Credit)
    End If
  Next

  'add cash acct number to fund numbers
  For cnt = 1 To TotalFunds
    FundCash$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(FundCash$) Then
      LSet AllFunds(cnt).FundCode = FundCash$ + SysCash$
    End If
  Next

  TotalGLIFS = NumOfWageAccts + GLIFArrayAdj + TotalFunds       '17

  ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14

  NextAcct = NumOfWageAccts + GLIFArrayAdj      '17

  For cnt = 1 To TotalFunds
    AcctNum$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(AcctNum$) Then
      NextAcct = NextAcct + 1
      GLIFRec(NextAcct).TranAcct = AcctNum$
      GLIFRec(NextAcct).TranDate = GLIFTDate$
      GLIFRec(NextAcct).TranDesc = "PR Net "
      GLIFRec(NextAcct).Source = GLIFSource$
      GLIFRec(NextAcct).FromFlag = "N"
      If AllFunds(cnt).Net > 0 Then
        GLIFRec(NextAcct).CrAmt = AllFunds(cnt).Net
        GLIFRec(NextAcct).DrAmt = 0
      ElseIf AllFunds(cnt).Net < 0 Then
        GLIFRec(NextAcct).DrAmt = Abs(AllFunds(cnt).Net)
        GLIFRec(NextAcct).CrAmt = 0
      End If
    End If
  Next

  'if using the imprest account then
  Select Case SysRec(1).USEIMP
  Case "I"      'was Y            'I C P
    TotalGLIFS = TotalGLIFS + 2
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    ReplaceString SysRec(1).ICRACCT, "-", ""
    GLIFRec(TotalGLIFS - 1).TranAcct = QPTrim$(SysRec(1).ICRACCT)
    GLIFRec(TotalGLIFS - 1).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS - 1).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS - 1).Source = GLIFSource$
    GLIFRec(TotalGLIFS - 1).FromFlag = "i"
    GLIFRec(TotalGLIFS - 1).DrAmt = Passed#(6)
    GLIFRec(TotalGLIFS - 1).CrAmt = 0
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "i"
    GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    GLIFRec(TotalGLIFS).DrAmt = 0
  Case "C"      'NEW Central Depository
    TotalGLIFS = TotalGLIFS + 1
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "c"

    If TotEIC# > 0 Then
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6) - TotEIC#
    Else
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    End If

    GLIFRec(TotalGLIFS).DrAmt = 0

    ReplaceString SysRec(1).ICRACCT, "-", ""
    ReDim CDGLIFRec(1 To TotalFunds) As GLIFDataType14
    For cnt = 1 To TotalFunds
      If AllFunds(cnt).Net <> 0 Then
        CDGLIFRec(cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + Left$(AllFunds(cnt).FundCode, 2)

        CDGLIFRec(cnt).TranDate = GLIFTDate$
        CDGLIFRec(cnt).TranDesc = "PPE " + GLIFTDate$
        CDGLIFRec(cnt).Source = GLIFSource$
        CDGLIFRec(cnt).FromFlag = "C"
        CDGLIFRec(cnt).DrAmt = AllFunds(cnt).Net                'Passed(6)
        CDGLIFRec(cnt).CrAmt = 0
      End If
    Next
  Case "P"

  End Select

  '*************END OF COMMON CODE SECTION

  If QPTrim$(SysRec(1).EXPMETHD) = "" Or SysRec(1).EXPMETHD = "0" Then
    GoTo WriteGLIFS
  End If

  ReDim SFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim MFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim RFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim SocGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim MedGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim RetGLIFRec(1 To NumOfWageAccts) As GLIFDataType14

  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)
  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)
  '
  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""

  For cnt = 1 To NumOfWageAccts
    'social
    SocGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(cnt).TranDate = GLIFTDate$
    SocGLIFRec(cnt).TranDesc = SOCMat$
    SocGLIFRec(cnt).Source = GLIFSource$
    SocGLIFRec(cnt).FromFlag = "S"
    SocGLIFRec(cnt).CrAmt = 0
    SocGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATSocAmt
    'medicare
    MedGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(cnt).TranDate = GLIFTDate$
    MedGLIFRec(cnt).TranDesc = MEDMat$
    MedGLIFRec(cnt).Source = GLIFSource$
    MedGLIFRec(cnt).FromFlag = "M"
    MedGLIFRec(cnt).CrAmt = 0
    MedGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATMedAmt
    'retirment
    RetGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(cnt).TranDate = GLIFTDate$
    RetGLIFRec(cnt).TranDesc = RETMat$
    RetGLIFRec(cnt).Source = GLIFSource$
    RetGLIFRec(cnt).FromFlag = "R"
    RetGLIFRec(cnt).CrAmt = 0
    RetGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATRetAmt
  Next

  For cnt = 1 To TotalFunds
    SFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + SOCLIAB$
    MFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + MEDLIAB$
    RFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + RETLIAB$
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(SFGLIFRec(Cnt2).TranAcct, 2) = Left$(SocGLIFRec(cnt).TranAcct, 2) Then
        SFGLIFRec(Cnt2).CrAmt = OldRound(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = SOCMatLib$
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      End If
    Next
  Next
  '
  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(MFGLIFRec(Cnt2).TranAcct, 2) = Left$(MedGLIFRec(cnt).TranAcct, 2) Then
        MFGLIFRec(Cnt2).CrAmt = OldRound(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(cnt).DrAmt)
        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = MEDMatLib$
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(RFGLIFRec(Cnt2).TranAcct, 2) = Left$(RetGLIFRec(cnt).TranAcct, 2) Then
        RFGLIFRec(Cnt2).CrAmt = OldRound(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = RETMatLib$
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      End If
    Next
  Next
  
WriteGLIFS:

  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  KillFile "TempIF.DAT"
  GHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GHandle Len = GLIFRecLen
  RecNo& = 1
  For cnt = 1 To TotalGLIFS
    If Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
'      If GLIFRec(cnt).DrAmt > 0 Or GLIFRec(cnt).CrAmt > 0 Then
      If GLIFRec(cnt).DrAmt <> 0 Or GLIFRec(cnt).CrAmt <> 0 Then '2/3/04 changed from line above
      'to allow for negative rounding adjustments in PCPrintPayRegisterSG under FixSplit
        Put GHandle, RecNo&, GLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(SocGLIFRec(cnt).TranAcct)) Then
      If SocGLIFRec(cnt).DrAmt > 0 Or SocGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, SocGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(MedGLIFRec(cnt).TranAcct)) Then
      If MedGLIFRec(cnt).DrAmt > 0 Or MedGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, MedGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(RetGLIFRec(cnt).TranAcct)) Then
      If RetGLIFRec(cnt).DrAmt > 0 Or RetGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, RetGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(SFGLIFRec(cnt).TranAcct)) Then
      If SFGLIFRec(cnt).DrAmt > 0 Or SFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, SFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(MFGLIFRec(cnt).TranAcct)) Then
      If MFGLIFRec(cnt).DrAmt > 0 Or MFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, MFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(RFGLIFRec(cnt).TranAcct)) Then
      If RFGLIFRec(cnt).DrAmt > 0 Or RFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, RFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  If SysRec(1).USEIMP = "C" Then
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(CDGLIFRec(cnt).TranAcct)) Then
        If CDGLIFRec(cnt).DrAmt <> 0 Or CDGLIFRec(cnt).CrAmt <> 0 Then
          Put GHandle, RecNo&, CDGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End If

  'added EIC GLIF records if present 6/07/94
  If TotEIC# > 0 Then
    For cnt = 1 To (SplDedCnt * 2)
      If EICGLIFRec(cnt).DrAmt <> 0 Or EICGLIFRec(cnt).CrAmt <> 0 Then
        Put GHandle, RecNo&, EICGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    Next
  End If
  
  Close GHandle
  
  ReDim GLSAmt(1) As String * 11
  ReDim GLMAmt(1) As String * 11

'**** NEW
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) Then
  If Exist(CurrCitiPath + "\" + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + "\" + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + "\" + GLAcctIdxFile) Then
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + "\" + GLAcctIdxFile
  Else
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + GLAcctIdxFile
  End If

  GLIdxRecLen = Len(JGLIdxRec(1))

  GLIFRecLen = Len(GLIFRec(1))

  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen

  If TotalGLIFS = 0 Then
    PrintGLRpt = False
    GoTo SkipGLRpt
  Else
    PrintGLRpt = True
  End If
  
  ReDim GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  GRHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GRHandle Len = GLIFRecLen
  
  For x = 1 To TotalGLIFS
    Get GRHandle, x, GLIFRec(x)
  Next x
'  SortG GLIFRec(), TotalGLIFS

  Do
    OutOfOrder = False                     'assume it's sorted
    For x = 1 To UBound(GLIFRec) - 1
      If GLIFRec(x).TranAcct > GLIFRec(x + 1).TranAcct Then
        SWAP GLIFRec(x), GLIFRec(x + 1)    'if we had to swap
        OutOfOrder = True                  'we're not done yet
      End If
    Next
  Loop While OutOfOrder

'04-05-96 Started Add summary to glif report
  For x = 1 To TotalGLIFS
    Put GRHandle, x, GLIFRec(x)
  Next x
  If TotalAccts = 0 Then
    Close
    GoTo SkipGLAccts
  End If
  Close GRHandle
  GLHandle = FreeFile
  Open GLIdxName$ For Random As GLHandle Len = GLIdxRecLen
'heredale
  ReDim JGoodAccts(1 To TotalAccts) As String * 16
  FrmShowPctComp.Label1 = "Printing Payroll Register"
  FrmShowPctComp.Show ' , Me
  For cnt = 1 To TotalAccts
    Get GLHandle, cnt, JGLIdxRec(1)
    ActualAccts = ActualAccts + 1
    ReplaceString JGLIdxRec(1).AcctNum, "-", ""
    JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
    FrmShowPctComp.ShowPctComp cnt, TotalAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close GLHandle
  TotalAccts = ActualAccts

SkipGLAccts:

'  Max = 55
'  Lines = 0
  TFRptName = "prrpts\GLFundSplTot.RPT"
  TFHandle = FreeFile
  Open TFRptName$ For Output As TFHandle
  
  RptName$ = "prrpts\PRGLIFG.RPT"
  KillFile RptName$
  GHandle = FreeFile
  Open RptName$ For Output As GHandle
'  RPTSetupPRN 14, GHandle
  THandle = FreeFile
  Open "TempIF.DAT" For Random As THandle Len = GLIFRecLen
'  GoSub GLIFHeader

'**** NEW
  FrmShowPctComp.Label1 = "Printing G\L Interface Report"
  FrmShowPctComp.Show

  For cnt = 1 To TotalGLIFS
    FrmShowPctComp.ShowPctComp cnt, TotalGLIFS
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If

    TotDebit# = OldRound(TotDebit# + GLIFRec(cnt).DrAmt)
    TotCredit# = OldRound(TotCredit# + GLIFRec(cnt).CrAmt)

    '10-13-94 ** added check for valid GL Account numbers
    If TotalAccts > 0 Then
      AcctOk = False
      For Cnt2 = 1 To TotalAccts
        If InStr(JGoodAccts(Cnt2), QPTrim$(GLIFRec(cnt).TranAcct)) And Len(QPTrim$(JGoodAccts(Cnt2))) = Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
          AcctOk = True
          Exit For
        End If
      Next
      If Not AcctOk Then
        GLError = -1
        LSet GLIFRec(cnt).Fill = "ERR     "
        Put THandle%, cnt, GLIFRec(cnt)
      Else
        LSet GLIFRec(cnt).Fill = "        "
      End If
    Else 'added this code during Windows port 3/22/2002
      GLIFRec(cnt).Fill = ""
    End If
    'NoCheckAccts:
    RSet GLSAmt(1) = Using(Image3$, GLIFRec(cnt).DrAmt)
    RSet GLMAmt(1) = Using(Image3$, GLIFRec(cnt).CrAmt)
    LSet EDesc(1) = QPTrim$(GLIFRec(cnt).TranDesc)
    NoAcctNum = Len(QPTrim$(GLIFRec(cnt).TranAcct))
    
    '------added 8/18/04-------------------------
    
    ThisDesc$ = QPTrim$(GLIFRec(cnt).TranDesc)
    If FundCount = 0 Then
      FundCount = FundCount + 1
      ReDim FundArray(1 To FundCount) As String
      FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
      ReDim FedTaxByFund(1 To FundCount) As Double
      ReDim EICByFund(1 To FundCount) As Double
      ReDim MedTaxByFund(1 To FundCount) As Double
      ReDim SocTaxByFund(1 To FundCount) As Double
      ReDim StaTaxByFund(1 To FundCount) As Double
      ReDim RetTaxByFund(1 To FundCount) As Double
      ReDim MedMatByFund(1 To FundCount) As Double
      ReDim SocMatByFund(1 To FundCount) As Double
      ReDim RetMatByFund(1 To FundCount) As Double
      ReDim DedAmtByFund(1 To 50, 1 To FundCount)
      Thisx = 1
   Else
     For x = 1 To FundCount
       If FundArray(x) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen) Then
         Thisx = x
         Exit For
       End If
     Next x
     If x > FundCount Then
       FundCount = FundCount + 1
       ReDim Preserve FundArray(1 To FundCount) As String
       FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
       ReDim Preserve FedTaxByFund(1 To FundCount) As Double
       ReDim Preserve EICByFund(1 To FundCount) As Double
       ReDim Preserve MedTaxByFund(1 To FundCount) As Double
       ReDim Preserve SocTaxByFund(1 To FundCount) As Double
       ReDim Preserve StaTaxByFund(1 To FundCount) As Double
       ReDim Preserve RetTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedMatByFund(1 To FundCount) As Double
       ReDim Preserve SocMatByFund(1 To FundCount) As Double
       ReDim Preserve RetMatByFund(1 To FundCount) As Double
       ReDim Preserve DedAmtByFund(1 To 50, 1 To FundCount)
       Thisx = FundCount
     End If
   End If

    If ThisDesc = "FED Withholdings" Then
      FedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "EIC Pmt" Then
      EICByFund(Thisx) = GLIFRec(cnt).DrAmt
    ElseIf ThisDesc = "SOC Withholdings" Then
      SocTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "MED Withholdings" Then
      MedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "STA Withholdings" Then
      StaTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "RET Withholdings" Then
      RetTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "SOC Match Liab" Then
      SocMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "MED Match Liab" Then
      MedMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "RET Match Liab" Then
      RetMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    Else
      For z = 1 To 50
        If ThisDesc = QPTrim$(DedCodes(z).DCDESC1) Then
          DedAmtByFund(z, Thisx) = GLIFRec(cnt).CrAmt
        End If
      Next z
    End If
    '----added 8/18/04--^^^^^^^^^^^
    
    If NoAcctNum > 0 Then
      If AcctOk = True Then
        LSet GLIFRec(cnt).Fill = " "
      End If
      Trans$ = QPTrim$(GLIFRec(cnt).TranAcct)
      LSet GLIFRec(cnt).TranAcct = " "
      LSet GLIFRec(cnt).TranAcct = Trans$
 '>>>>>>>>>>>>>>>>MAIN PRINT LINE<<<<<<<<<<<<<<<<<<
      Print #GHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm; GLIFRec(cnt).TranAcct; dlm; GLIFRec(cnt).Fill; dlm; EDesc(1); dlm; GLSAmt(1); dlm; GLMAmt(1); dlm; TotalAccts
      Lines = Lines + 1
    End If
  Next cnt
  GoSub GLIFTotals
  Fund$ = Left$(QPTrim$(GLIFRec(1).TranAcct), 2)
  
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
    If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
      CIDebit = GLIFRec(1).DrAmt
      CICredit = GLIFRec(1).CrAmt
      GoTo NotNow1
    End If
  End If
  
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  
NotNow1:
  For cnt = 2 To TotalGLIFS
    NFund$ = Left$(QPTrim$(GLIFRec(cnt).TranAcct), 2)
    If NFund$ <> Fund$ Then
      GoSub PrintFundTotal
      Fund$ = NFund$
      If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
        If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
          CIDebit = GLIFRec(cnt).DrAmt
          CICredit = GLIFRec(cnt).CrAmt
          GoTo NotNow2
        End If
      End If
      FDebit# = GLIFRec(cnt).DrAmt
      FCredit# = GLIFRec(cnt).CrAmt
    Else
      If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
        If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
          CIDebit# = OldRound(CIDebit# + GLIFRec(cnt).DrAmt)
          CICredit# = OldRound(CICredit# + GLIFRec(cnt).CrAmt)
          GoTo NotNow2
        End If
      End If
      FDebit# = OldRound(FDebit# + GLIFRec(cnt).DrAmt)
      FCredit# = OldRound(FCredit# + GLIFRec(cnt).CrAmt)
    End If
NotNow2:
  Next

  GoSub PrintFundTotal

  Close GHandle
  Close TFHandle
  Close THandle
  
  If GLError <> -1 Then
    GLError = TotalAccts
  End If

SkipGLRpt:
  
  Exit Sub

PrintFundTotal: '8/10/04
  For x = 1 To FundCount
    If FundArray(x) = Fund Then
      Thisx = x
      Exit For
    End If
  Next x

'  GLDebitTotal = 0
'  GLCreditTotal = 0
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
    If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
      GLSAmt(1) = CIDebit
      GLMAmt(1) = CICredit
      GoTo CenImp
    End If
  End If
  
  GLSAmt(1) = FDebit#
  GLMAmt(1) = FCredit#
  
CenImp:
  GLIFRec(1).TranAcct = Fund$
  '                                    0                       1              2
  Print #TFHandle, GLIFRec(1).TranAcct; dlm; GLSAmt(1); dlm; GLMAmt(1); dlm;
  '8/17/04 5/6/09
  TOTFEDTAX = OldRound(TOTFEDTAX + FedTaxByFund(Thisx))
  TOTEICFORRPT = OldRound(TOTEICFORRPT + EICByFund(Thisx))
  TOTMEDTAX = OldRound(TOTMEDTAX + MedTaxByFund(Thisx))
  TOTSOCTAX = OldRound(TOTSOCTAX + SocTaxByFund(Thisx))
  TOTSTATAX = OldRound(TOTSTATAX + StaTaxByFund(Thisx))
  TOTRetTax = OldRound(TOTRetTax + RetTaxByFund(Thisx))
  TOTMEDMat = OldRound(TOTMEDMat + MedMatByFund(Thisx))
  TOTSOCMat = OldRound(TOTSOCMat + SocMatByFund(Thisx))
  TOTRETMat = OldRound(TOTRETMat + RetMatByFund(Thisx))
  '                       3                         4                        5                            6                         7
  Print #TFHandle, FedTaxByFund(Thisx); dlm; MedTaxByFund(Thisx); dlm; SocTaxByFund(Thisx); dlm; StaTaxByFund(Thisx); dlm; RetTaxByFund(Thisx); dlm;
  Print #TFHandle, MedMatByFund(Thisx); dlm; SocMatByFund(Thisx); dlm; RetMatByFund(Thisx); dlm;
  For z = 1 To 50
    TotDedAmt(z) = OldRound(TotDedAmt(z) + DedAmtByFund(z, Thisx))
    If DedAmtByFund(z, Thisx) > 0 Then
      '                    11                  12
      Print #TFHandle, QPTrim$(DedCodes(z).DCDESC1); dlm; DedAmtByFund(z, Thisx); dlm;
    Else
      TotDedAmt(z) = TotDedAmt(z) + 0
      Print #TFHandle, QPTrim$(DedCodes(z).DCDESC1); dlm; 0; dlm;
    End If
  Next z
  Print #TFHandle, NumOfDeds; dlm;
  
  Print #TFHandle, TOTFEDTAX; dlm; TOTMEDTAX; dlm; TOTSOCTAX; dlm; TOTSTATAX; dlm;
  '
  Print #TFHandle, TOTRetTax; dlm; TOTMEDMat; dlm; TOTSOCMat; dlm; TOTRETMat; dlm;
  '
  For z = 1 To 50
    If TotDedAmt(z) > 0 Then
      Print #TFHandle, QPTrim$(DedCodes(z).DCDESC1); dlm; TotDedAmt(z); dlm;
    Else
      Print #TFHandle, QPTrim$(DedCodes(z).DCDESC1); dlm; 0; dlm;
    End If
  Next z
  
  Print #TFHandle, TotDebit; dlm; TotCredit; dlm; -EICByFund(Thisx); dlm; -TOTEICFORRPT
  
  GLDebitTotal = OldRound(GLDebitTotal + GLSAmt(1))
  GLCreditTotal = OldRound(GLCreditTotal + GLMAmt(1))
Return

GLIFTotals:
  Print #GHandle, QPTrim$(Unit(1).UFEMPR); dlm; MakeRegDate(PDR(1).PEREND); dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotalAccts
'  Print #GHandle,
'  Print #GHandle, Dash0(1)
'  Print #GHandle, "   Totals:                                  " + Using(Image3$, TotDebit#) + "     " + Using(Image3$, TotCredit#)
'  If TotalAccts = 0 Then
'    Print #TFHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory."; dlm; "" '+ CrLf$
'  Else
'    Print #TFHandle, Using(Image3$, TotDebit#); dlm; Using(Image3$, TotCredit#)
'  End If
'  If TotalAccts = 0 Then
'    Print #GHandle,
'    Print #GHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory." '+ CrLf$
'    Print #TFHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory." '+ CrLf$
'  End If
'  Print #GHandle,
'  Print #GHandle, "SubTotals:"
Return

'GLIFHeader:
'  Print #GHandle, QPTrim$(Unit(1).UFEMPR)
'  Print #GHandle, "G/L Interface Report."
'  Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
'  Print #GHandle,
'  Print #GHandle, "Account No.           Description                Debit         Credit"
'
'  Print #GHandle, Dash0(1)
'  Lines = 6
'Return

End Sub
Private Sub SortT(DistbSumAccts() As DistWageRptType, NumOfWageAccts As Integer)

  Dim TempDAccts As DistWageRptType
  Dim Temp As String
  Dim Largest As Double
  Dim Smallest As Double
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  ReDim OrderedList(1 To NumOfWageAccts) As DistWageRptType
  ReDim IdxDAccts(1 To NumOfWageAccts) As Integer
  
  Largest = 0
  For x = 1 To NumOfWageAccts
    thisNum = ReplaceString(DistbSumAccts(x).Acct, "-", "")
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To NumOfWageAccts
    thisNum = ReplaceString(DistbSumAccts(x).Acct, "-", "")
     If Val(thisNum) < Smallest Then
         Smallest = Val(thisNum)
         Thisx = x
      End If
    Next x
    OrderedList(Nextx) = DistbSumAccts(Thisx)
    TempDAccts = DistbSumAccts(Nextx)
    DistbSumAccts(Nextx) = DistbSumAccts(Thisx)
    DistbSumAccts(Thisx) = TempDAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = NumOfWageAccts + 1
End Sub

Private Sub SortE(SplitDeds() As SplitDedsType, NumOfDeds As Integer)
  Dim TempDAccts As SplitDedsType
  Dim Temp As String
  Dim Largest As Single
  Dim Smallest As Single
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  ReDim OrderedList(1 To NumOfDeds) As SplitDedsType
  ReDim IdxDAccts(1 To NumOfDeds) As Integer
  
  Largest = 0
  For x = 1 To NumOfDeds
    thisNum = ReplaceString(SplitDeds(x).Acct, "-", "")
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To NumOfDeds
     thisNum = ReplaceString(SplitDeds(x).Acct, "-", "")
      If Val(thisNum) < Smallest Then
        Smallest = Val(thisNum)
        Thisx = x
      End If
    Next x
    OrderedList(Nextx) = SplitDeds(Thisx)
    TempDAccts = SplitDeds(Nextx)
    SplitDeds(Nextx) = SplitDeds(Thisx)
    SplitDeds(Thisx) = TempDAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = NumOfDeds + 1

End Sub

Private Sub SWAP(This As GLIFDataType14, ForThis As GLIFDataType14)
  Dim TempThis As GLIFDataType14
  TempThis = This
  This = ForThis
  ForThis = TempThis
End Sub

Private Sub SortF(AllFunds() As FundType, TotalFunds As Integer)
  Dim TempFAccts As FundType
  Dim Temp As String
  Dim Largest As Single
  Dim Smallest As Single
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  ReDim OrderedList(1 To TotalFunds) As FundType
  ReDim IdxFAccts(1 To TotalFunds) As Integer
  
  Largest = 0
  For x = 1 To TotalFunds
    thisNum = ReplaceString(AllFunds(x).FundCode, "-", "")
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To TotalFunds
    thisNum = ReplaceString(AllFunds(x).FundCode, "-", "")
      If Val(thisNum) < Smallest Then
         Smallest = Val(thisNum)
         Thisx = x
      End If
    Next x
    OrderedList(Nextx) = AllFunds(Thisx)
    TempFAccts = AllFunds(Nextx)
    AllFunds(Nextx) = AllFunds(Thisx)
    AllFunds(Thisx) = TempFAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = TotalFunds + 1

End Sub

Private Sub SortG(GLIFRec() As GLIFDataType14, TotalGLIFS As Integer)
  Dim TempGAccts As GLIFDataType14
  Dim Temp As String
  Dim Largest As Single
  Dim Smallest As Single
  Dim thisNum As String
  Dim x As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  ReDim OrderedList(1 To TotalGLIFS) As GLIFDataType14
  ReDim IdxFAccts(1 To TotalGLIFS) As Integer
  
  Largest = 0
  For x = 1 To TotalGLIFS
    thisNum = ReplaceString(GLIFRec(x).TranAcct, "-", "")
    If Val(thisNum) > Largest Then
      Largest = Val(thisNum)
    End If
  Next x
  
  Smallest = Largest + 1
  Nextx = 1
  Thisx = Nextx
  Do
    For x = Nextx To TotalGLIFS
    thisNum = ReplaceString(GLIFRec(x).TranAcct, "-", "")
    If Val(thisNum) < Smallest Then
       Smallest = Val(thisNum)
       Thisx = x
    End If
    Next x
    OrderedList(Nextx) = GLIFRec(Thisx)
    TempGAccts = GLIFRec(Nextx)
    GLIFRec(Nextx) = GLIFRec(Thisx)
    GLIFRec(Thisx) = TempGAccts
    Nextx = Nextx + 1
    Smallest = Largest + 1
  Loop Until Nextx = TotalGLIFS + 1

End Sub

Sub PCPrintPayRegisterST(PathCode As Integer)
  
  Dim TransRec(1) As TransRecType
  Dim EmpRec1(1) As EmpData1Type
  Dim PDR(1) As PeriodDefaultRecType
  Dim Unit(1) As UnitFileRecType

  ReDim DistbSumAccts(1 To 1) As DistWageRptType
  Dim SysRec(1) As RegDSysFileRecType
  Dim GLIFRec(1 To 1) As GLIFDataType12
  ReDim SplitDeds(1 To 1) As SplitDedsType
  ReDim EmpSpl(1 To 1) As SplitDedsType
  ReDim EmpAct(1 To 1) As DistWageRptType
  Dim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  Dim ErnCodes(1 To 3) As ErnCodeRecType
  Dim TotDeds#(1 To 50), PrintGLRpt As Boolean
  Dim TotErns(1 To 3) As Double
  
  Dim EDesc(1) As String * 21
  Dim EDAct(1) As String * 14
  Dim EDPct(1) As String * 11
  Dim EDRHrs(1) As String * 11
  Dim EDOHrs(1) As String * 11
  Dim EDRPay(1) As String * 11
  Dim EDOPay(1) As String * 11
  Dim EDEarn(1) As String * 11
  Dim EDGroP(1) As String * 11
  Dim EDSAmt(1) As String * 11
  Dim EDMAmt(1) As String * 11
  Dim EDRAmt(1) As String * 11
  Dim ENumb(1) As String * 14 'was ENum
  Dim EName(1) As String * 33

  Dim BRat(1) As String * 11
  Dim ORat(1) As String * 11
  Dim TaxFrn(1) As String * 11
  Dim Fill11(1) As String * 11
  Dim SCnt(1) As String * 11
  Dim HCnt(1) As String * 11
  Dim RHrs(1) As String * 11
  Dim VHrs(1) As String * 11
  Dim SHrs(1) As String * 11
  Dim HHrs(1) As String * 11
  Dim CHrs(1) As String * 11

  Dim THrs(1) As String * 11
  Dim PHrs(1) As String * 11
  Dim OTHrs(1) As String * 11
  Dim OTPaid(1) As String * 11
  Dim OTComp(1) As String * 11
  Dim RErnP(1) As String * 11
  Dim OErnP(1) As String * 11
  Dim GPayP(1) As String * 11
  Dim SSTaxP(1) As String * 11
  Dim MTaxP(1) As String * 11
  Dim FTaxP(1) As String * 11
  Dim STaxP(1) As String * 11
  Dim RetirP(1) As String * 11
  Dim NetPayP(1) As String * 11
  Dim Ded(1) As String * 11

  'added for EIC   6/07/94
  Dim EEicP(1) As String * 11
  Dim Ern(1) As String * 11
  Dim Pg(1) As String * 5
  Dim EMPLine(1) As String * 132
  Dim SplDed(1) As String * 7
  Dim SplDed2(1) As String * 9
  Dim Dash0(1) As String * 69
  Dim Dash(1) As String * 132
  Dim Dash2(1) As String * 124
  Dim Dash3(1) As String * 91
  Dim PPDefaultHandle As Integer
  Dim FileHandle As Integer
  Dim DedCodeFileHandle As Integer
  Dim ErnCodeFileHandle As Integer
  Dim Emp1Handle As Integer
  Dim x As Integer, LincCnt As Integer
  Dim EmpIdxNNameHandle As Integer
  Dim NumIdxRec As NumbSortIdxType
  Dim FACnt As Integer, RetCode As Integer
  Dim Image0$, Image$, Image3$, Image4$, Image5$
  Dim GFedGross#, GStaGross#, GMedGross#, GSocGross#
  Dim GRetGross#, GTaxFrn#, TotDebit#, TotCredit#
  Dim ArraySize As Integer, EmpActLen As Integer
  Dim DistbSumAcctsLen As Integer, TransRecLen As Long
  Dim Emp1RecLen As Integer, IdxRecLen As Integer
  Dim IdxFileSize&, NumOfRecs As Long, SalCnt As Integer
  Dim HrlCnt As Integer, DLineCnt As Integer
  Dim LineCnt As Integer, NumOfWageAccts As Integer
  Dim MaxLines As Integer, DMaxLines As Integer
  Dim EPage As Integer, Page As Integer
  Dim STitle$(1 To 5), TDed$, DTitle$(1 To 5), AMAmt#
  Dim FTitle$
  Dim LastDed As Integer, DCnt As Integer
  Dim ETitle$, TErn$, LastErn As Integer, SumHeader2$
  Dim PayRegisterRptName$, RHandle As Integer
  Dim DHandle As Integer, DistributionRptName$
  Dim NHandle As Integer, THandle As Integer
  Dim Cnt2 As Integer, cnt As Integer
  Dim Acct$, ENumOfAct As Integer, First As Integer
  Dim Start As Integer, Last As Integer
  Dim TotalSocAmt#, LastActive As Integer
  Dim Cnt3 As Integer, DistDif#, TotalMedAmt#
  Dim TotalRetAmt#, DLincCnt As Integer
  Dim Passed#(1 To 6), CCnt As Integer
  Dim FTax#, STax#, SSTax#, MTax#, RETTOT#, TNetPay#
  Dim RptTitle$, RegHrs#, VACHRS#, SICKHRS#, HOLHRS#
  Dim COMPHRS#, PerHRS#, TotalHrs#, TotEIC#, TotHrs#
  Dim TOTPaid#, TOTComp#, TRegWage#, TOTWage#, GPay#
  Dim SumDed$(1 To 5), SysHandle As Integer, DedAmt#
  Dim SumErn$, FF$, ARAmt#, CCnt3 As Integer
  Dim RegWage#, OTWage#, AddEarn#, DGPay#, ASAmt#
  Dim SplDedCnt As Integer, TempLine$(1 To 50), SFund$, zCnt As Integer
  Dim EmpSplCnt As Integer, ESCnt As Integer, SPct@
  Dim ESplAmt#, GotIt As Boolean, CCnt01 As Integer, SDedAmt#
  Dim tFedTax#, tStaTax#, tMedTax#, aa As Integer
  Dim tSocTax#, tRetTax#, tEICPay#
  Dim Nextx As Integer
  Dim tripCnt As Integer
  Dim NumOfDeds As Integer
  Dim DedDescTrimmed(1 To 50) As String * 8
  Dim TempMedTaxAmt#
  
  Dim CSocWHAcct$
  Dim CMedWHAcct$
  Dim CSocMatchAcct$
  Dim CMedMatchAcct$
  Dim CRetMatchAcct$
  Dim CFedWHAcct$
  Dim CStaWHAcct$
  Dim CRetWHAcct$
  Dim CDedAcct$
  Dim CPRNetAcct$
  Dim DWagesAcct$
  Dim DSocMatchAcct$
  Dim DMedMatchAcct$
  Dim DRetMatchAcct$
  Dim FundNumOnly$
  Dim FundAndAcctOnly$
  Dim TempVoid As VoidCheckType
  Dim TVHandle As Integer
  Dim TVCnt As Double
'  Dim FundLen As Integer
  Dim AcctLen As Integer
  Dim DetLen As Integer
  Dim ThisPR As Double
  Dim NextSpl As Integer
  Dim ActiveCnt As Integer
  
  ActiveCnt = 0
  FF$ = Chr$(12)
  OpenPPDefaultFile PPDefaultHandle
  Get PPDefaultHandle, 1, PDR(1)
  Close PPDefaultHandle
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle

  OpenDedCodeFile DedCodeFileHandle
  ReDim DedCodeNums(1 To 1) As String '6/15/04
  For x = 1 To 50
    Get DedCodeFileHandle, x, DedRec
    DedCodes(x) = DedRec
    RSet DedDescTrimmed(x) = QPTrim$(DedRec.DCDESC1)
    If Len(QPTrim(DedRec.DCDESC1)) > 0 Then
      NumOfDeds = NumOfDeds + 1
      ReDim Preserve DedCodeNums(1 To NumOfDeds) As String '6/15/04
      DedCodeNums(NumOfDeds) = QPTrim$(DedRec.DCACCT1)
    End If
  Next x
  Close DedCodeFileHandle

  OpenErnCodeFile ErnCodeFileHandle
  For x = 1 To 3
    Get ErnCodeFileHandle, x, ErnCodes(x)
  Next x
  Close ErnCodeFileHandle
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec(1)
  Close SysHandle
  FACnt = SysRec(1).AcctCnt
  
'  Call GetAcctStruct(SysRec(1).CITIDIR, FundLen, AcctLen, DetLen)
  Call GetAcctStruct(CurrCitiPath, FundLen, AcctLen, DetLen)
  
  Image0$ = "###0"
  Image$ = "###0.00"
  Image3$ = "###,##0.00"
  Image4$ = "###.####"
  Image5$ = "###,##0.00"
  
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0
  GTaxFrn# = 0
  TotDebit# = 0
  TotCredit# = 0
  
  ArraySize = 1
  
  LSet Fill11$(1) = ""
  LSet Dash0(1) = String$(69, "-")
  LSet Dash(1) = String$(132, "-")
  LSet Dash2(1) = String$(124, "-")
  RSet Dash3(1) = String$(63, "-")
  
  EmpActLen = Len(EmpAct(1))
  DistbSumAcctsLen = Len(DistbSumAccts(1))
  
  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))
  
  OpenEmpData1File Emp1Handle
  NumOfRecs = LOF(Emp1Handle) / Len(EmpRec1(1))
  Close Emp1Handle
  
  SalCnt = 0
  HrlCnt = 0
  
  DLineCnt = 0
  LineCnt = 0
  
  NumOfWageAccts = 0
  
  MaxLines = 45
  DMaxLines = 45
  EPage = 1

  Page = 1
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  For x = 1 To NumOfRecs
    Get EmpIdxNNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxNNameHandle
  
  For x = 1 To 5
    STitle$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For cnt = 1 To NumOfDeds
    TDed$ = Left$((QPTrim$(DedCodes(cnt).DCDESC1) + "       "), 6)
    If Len(QPTrim(TDed$)) = 0 Then GoTo SkipIt1
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    STitle$(Nextx) = STitle$(Nextx) + TDed$ + " "
    tripCnt = tripCnt + 1
SkipIt1:
  Next
  STitle$(1) = "Fund  " + STitle$(1) + "   Fed    Sta    Med    Soc    Ret"
  STitle$(2) = "      " + STitle$(2)
  STitle$(3) = "      " + STitle$(3)
  STitle$(4) = "      " + STitle$(4)
  STitle$(5) = "      " + STitle$(5)
  
  For x = 1 To 5
    DTitle$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For cnt = 1 To NumOfDeds
    TDed$ = QPTrim$(DedCodes(cnt).DCDESC1)
    If Len(QPTrim(TDed$)) = 0 Then GoTo SkipIt2
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    LastDed = LastDed + 1
    RSet Ded(1) = TDed$
    DTitle$(Nextx) = DTitle$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
SkipIt2:
  Next cnt
  
  '---------------------------------------------
  
  ETitle$ = ""
  For cnt = 1 To 3
    TErn$ = QPTrim$(ErnCodes(cnt).ERNCODE1)
    If Len(TErn$) > 0 Then
      LastErn = LastErn + 1
      RSet Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    Else
      Exit For
    End If
  Next
  
  If LastErn < 3 Then
    ETitle$ = Space$(11 * (3 - LastErn)) + ETitle$
  End If
  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  
  '------------------------------------------------------------------
  PayRegisterRptName$ = "PRRPTS\REGISTER.RPT"
  RHandle = FreeFile
  Open PayRegisterRptName$ For Output As RHandle
  RPTSetupPRN 13, RHandle
  DistributionRptName$ = "PRRPTS\DISTRIBU.RPT"
  DHandle = FreeFile
  Open DistributionRptName$ For Output As DHandle
  RPTSetupPRN 13, DHandle
  
  OpenEmpData1File NHandle
  OpenTransWorkFile THandle
  
  GoSub PrintPayRollHeader
  GoSub PrintDistHeader
  FrmShowPctComp.Label1 = "Printing Earnings Distribution Register"
  FrmShowPctComp.Show ' , Me
  
  KillFile TempVoidFileName
  OpenTempVoidFile TVHandle
  
  For cnt = 1 To NumOfRecs
    Get THandle, IdxBuff(cnt), TransRec(1)
    If TransRec(1).TActive = True Then
      Get NHandle, IdxBuff(cnt), EmpRec1(1)
      GoSub SumAndPrintTime
      GoSub ParseDistributions
      ActiveCnt = ActiveCnt + 1
      LineCnt = LineCnt + 5 + (NumOfDeds \ 12) 'added 5/28
      If LineCnt >= MaxLines And cnt < NumOfRecs Then
        LineCnt = 0
        Print #RHandle, FF$
        GoSub PrintPayRollHeader
      End If
    End If
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
    
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
BadEmp:
  Next
  
  If ActiveCnt = 0 Then
    MsgBox "No employees have been designated for payroll processing."
    Close
    GoTo AltExit
  End If
  
  GoSub PrintSumTotal
  GoSub PrintDistTotal
  GoSub PrintSplitTotal

  Close THandle
  Close NHandle
  RPTSetupPRN 123, RHandle '7/24
  Close RHandle
  RPTSetupPRN 123, DHandle '7/24
  Close DHandle
  'HERE:  register report files have been written to disk
  'GLIF START
  'if there is a GL transfer directory then make GLIF file
  
'  If Len(QPTrim$(SysRec(1).CITIDIR)) > 0 Then
  If Len(CurrCitiPath) > 0 Then
    Passed#(1) = FTax#
    Passed#(2) = STax#
    Passed#(3) = SSTax#
    Passed#(4) = MTax#
    Passed#(5) = RETTOT#
    Passed#(6) = TNetPay#
  
    MakeGLIFFileST SplitDeds(), DistbSumAccts(), TotEIC#, TotDeds#(), Passed#(), PrintGLRpt
    'unrem
  End If
  
  Unload frmLoadingRpt
  RptTitle$ = "Payroll Register Report"
  ViewPrint PayRegisterRptName, RptTitle$, True
  
  If RetCode = -1 Then GoTo AltExit
  
  RptTitle$ = "Payroll Distribution Report"
  ViewPrint DistributionRptName, RptTitle$, True
  
  If RetCode = -1 Then GoTo AltExit
  
  If PrintGLRpt Then
    If GLCreditTotal <> GLDebitTotal Then
      frmMessage.Label1.Caption = "The General Ledger Interface is OUT OF BALANCE."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      MainLog "User warned that the GL Interface is OUT OF BALANCE (Debit Total = " + QPTrim$(Using$("$#,###,##0.00", GLDebitTotal)) + " and the Credit Total is " + QPTrim$(Using$("$#,###,##0.00", GLCreditTotal)) + ")."
    End If
    If GLError < 1 Then
      frmMessage.Label1.Caption = "General Ledger number errors have been found in the GL Interface Report."
      frmMessage.Label1.Top = 800
      frmMessage.Show vbModal
      MainLog "User warned that the GL Interface Report has GL number errors."
    End If
    RptTitle$ = " Report"
    ViewPrint "prrpts\PRGLIF.RPT", RptTitle$, True
  End If
  
AltExit:
  
Exit Sub
  
PrintPayRollHeader:
  RSet Pg(1) = Str$(Page)
  Print #RHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1) ' + CrLf$
  Print #RHandle, "Payroll Register" '+ CrLf$
  Print #RHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$ + CrLf$
  Print #RHandle, "Employee No   Name                                                                                                               EIC" ' + CrLf$
  Print #RHandle, "  Base Rate   O/T Rate   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid   O/T Comp" ' + CrLf$
  Print #RHandle, ETitle$ '+ CrLf$ + DTitle$ + CrLf$ + Dash(1) ' + CrLf$
  Dim AddLine As Integer
  For x = 1 To 5
    If Len(QPTrim(DTitle(x))) > 0 Then
      AddLine = AddLine + 1
      Print #RHandle, DTitle$(x)
    End If
  Next x
  Print #RHandle, Dash(1)
  LincCnt = LineCnt + AddLine
  Page = Page + 1
Return
  
SumAndPrintTime:
  RegHrs# = OldRound#(RegHrs# + TransRec(1).RegHrsWork)
  VACHRS# = OldRound#(VACHRS# + TransRec(1).VacUsed)
  SICKHRS# = OldRound#(SICKHRS# + TransRec(1).SickUsed)
  HOLHRS# = OldRound#(HOLHRS# + TransRec(1).HOLHOURS)
  COMPHRS# = OldRound#(COMPHRS# + TransRec(1).CompUsed)
  PerHRS# = OldRound#(PerHRS# + TransRec(1).PerHours)
  
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  TotalHrs# = OldRound(TotalHrs# + TransRec(1).PerHours)

  'added for EIC   6/07/94
  TotEIC# = OldRound#(TotEIC# + TransRec(1).EICAmt)
  
  '-=-=-=-=-=-=-=
  TotHrs# = OldRound#(TotHrs# + TransRec(1).OTHours)
  If TransRec(1).OTHrsPaid > 0 Then
    TOTPaid# = OldRound#(TOTPaid# + TransRec(1).OTHrsPaid)
  End If
  TOTComp# = OldRound#(TOTComp# + TransRec(1).OT2Comp)
  
  TRegWage# = OldRound#(TRegWage# + TransRec(1).TotRegWage)
  
  If TransRec(1).TotOTWage > 0 Then
    TOTWage# = OldRound#(TOTWage# + TransRec(1).TotOTWage)
  End If
  
  GPay# = OldRound#(GPay# + TransRec(1).GrossPay)
  SSTax# = OldRound#(SSTax# + TransRec(1).SocTaxAmt)
  MTax# = OldRound#(MTax# + TransRec(1).MedTaxAmt)
  FTax# = OldRound#(FTax# + TransRec(1).FedTaxAmt)
  STax# = OldRound#(STax# + TransRec(1).StaTaxAmt)
  If TransRec(1).RetireAmt > 0 Then
    RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
  End If
  
  TNetPay# = OldRound#(TNetPay# + TransRec(1).NetPay)
  GFedGross# = OldRound#(GFedGross# + TransRec(1).FedGrossPay)
  GStaGross# = OldRound#(GStaGross# + TransRec(1).StaGrossPay)
  GSocGross# = OldRound#(GSocGross# + TransRec(1).SocGrossPay)
  GRetGross# = OldRound#(GRetGross# + TransRec(1).RetGrossPay)
  GMedGross# = OldRound#(GMedGross# + TransRec(1).MedGrossPay)
  
  GTaxFrn# = OldRound#(GTaxFrn# + TransRec(1).TaxFring)
  
  LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
  LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)
  RSet BRat(1) = Using(Image3$, TransRec(1).BaseRate)
  RSet ORat(1) = Using(Image3$, TransRec(1).OTRate)
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsWork)
  RSet VHrs(1) = Using(Image$, TransRec(1).VacUsed)
  RSet SHrs(1) = Using(Image$, TransRec(1).SickUsed)
  RSet HHrs(1) = Using(Image$, TransRec(1).HOLHOURS)
  RSet CHrs(1) = Using(Image$, TransRec(1).CompUsed)
  RSet THrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  RSet TaxFrn(1) = Using(Image$, TransRec(1).TaxFring)
  RSet OTHrs(1) = Using(Image$, TransRec(1).OTHours)
  RSet PHrs(1) = Using(Image$, TransRec(1).PerHours)

  RSet OTPaid(1) = Using(Image$, TransRec(1).OTHrsPaid)
  RSet OTComp(1) = Using(Image$, TransRec(1).OT2Comp)
  
  'added for EIC     6/07/94
  RSet EEicP(1) = Using(Image3$, TransRec(1).EICAmt)
  
  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  '=======
  RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
  RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
  
  RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
  RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
  RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
  RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
  RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
  RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
  RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    TotDeds#(Cnt2) = OldRound#(TotDeds#(Cnt2) + TransRec(1).DAmt(Cnt2))
    RSet Ded(1) = Using(Image3$, TransRec(1).DAmt(Cnt2))
    tripCnt = tripCnt + 1
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
  Next
  
  '----------------------------------------------
  SumErn$ = ""
'  For Cnt2 = 1 To LastErn
  For Cnt2 = 1 To 3
    If Cnt2 > LastErn Then
      Select Case LastErn
      Case 2:
        SumErn$ = "           " + SumErn$
        Exit For
      Case 1:
        SumErn$ = "           " + "           " + SumErn$
        Exit For
      Case Else:
        Exit For
      End Select
    End If
    TotErns(Cnt2) = OldRound#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSet Ern(1) = Using(Image3$, TransRec(1).EAmt(Cnt2))
NoMoreErns:
    SumErn$ = SumErn$ + Ern(1)
  Next
  '-------------------------------------------------------
  RSet EMPLine$(1) = EEicP(1)
  Mid$(EMPLine$(1), 1) = ENumb(1) + EName(1)
  Print #RHandle, EMPLine$(1)
  Print #RHandle, BRat(1) + ORat(1) + TaxFrn(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + OTComp(1) ' + CrLf$
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1)  ' + CrLf$
  For x = 1 To 5
    If Len(QPTrim$(SumDed(x))) > 0 Then
      Print #RHandle, SumDed$(x)
    End If
  Next x
  Print #RHandle,
Return
  
PrintSumTotal:
  RSet SCnt(1) = Using(Image0$, SalCnt)
  RSet HCnt(1) = Using(Image0$, HrlCnt)
  
  RSet Fill11(1) = Using(Image3$, GTaxFrn#)
  RSet THrs(1) = Using(Image3$, TotalHrs#)
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet VHrs(1) = Using(Image3$, VACHRS#)
  RSet SHrs(1) = Using(Image3$, SICKHRS#)
  RSet HHrs(1) = Using(Image3$, HOLHRS#)
  RSet CHrs(1) = Using(Image3$, COMPHRS#)
  RSet PHrs(1) = Using(Image3$, PerHRS#)
  
  RSet OTPaid(1) = Using(Image3$, TOTPaid#)
  RSet OTComp(1) = Using(Image3$, TOTComp#)
  
  RSet RErnP(1) = Using(Image3$, TRegWage#)
  RSet OErnP(1) = Using(Image3$, TOTWage#)
  
  RSet GPayP(1) = Using(Image3$, GPay#)
  RSet SSTaxP(1) = Using(Image3$, SSTax#)
  RSet MTaxP(1) = Using(Image3$, MTax#)
  RSet FTaxP(1) = Using(Image3$, FTax#)
  RSet STaxP(1) = Using(Image3$, STax#)
  RSet RetirP(1) = Using(Image3$, RETTOT#)
  RSet NetPayP(1) = Using(Image3$, TNetPay#)
  
  'added for EIC      6/07/94
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds
      If tripCnt = 13 Then
        tripCnt = 1
        Nextx = Nextx + 1
      End If
      RSet Ded(1) = Using(Image3$, TotDeds#(Cnt2))
      SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
      tripCnt = tripCnt + 1
  Next
  '---------------------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, TotErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  
  '--------------NEW----------------------------
  RSet Pg(1) = Str$(Page)
  Print #RHandle, FF$
  Print #RHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1) ' + CrLf$
  Print #RHandle, "Payroll Register Summary" ' + CrLf$
  Print #RHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
  Print #RHandle, 'CrLf$ + Dash(1) + CrLf$ + CrLf$
  Print #RHandle,
  Print #RHandle, Dash(1)
  Print #RHandle,
  Print #RHandle, "   Salaried     Hourly   Tax Frng    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid   O/T Comp" ' + CrLf$
  Print #RHandle, SCnt(1) + HCnt(1) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + OTComp(1) ' + CrLf$
  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"
  
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1) ' + CrLf$
  Print #RHandle,
  For x = 1 To 5
    If Len(QPTrim(DTitle$(x))) > 0 Then
      Print #RHandle, DTitle$(x)
      Print #RHandle, SumDed$(x)
    End If
  Next x
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Med Gross  Soc Gross  Ret Gross  EIC Total" ' + CrLf$

  RSet FTaxP(1) = Using(Image5$, GFedGross#)
  RSet STaxP(1) = Using(Image5$, GStaGross#)
  RSet MTaxP(1) = Using(Image5$, GMedGross#)
  RSet SSTaxP(1) = Using(Image5$, GSocGross#)
  RSet RetirP(1) = Using(Image5$, GRetGross#)
  RSet EEicP(1) = Using(Image5$, TotEIC#)
  
  Print #RHandle, FTaxP(1) + STaxP(1) + MTaxP(1) + SSTaxP(1) + RetirP(1) + EEicP(1) ' + CrLf$
  Print #RHandle,
  Print #RHandle, Dash(1)
  Print #RHandle, FF$
Return
  
  
  '-----------------------------------------------------------------------
ParseDistributions:

  ReDim EmpAct(1 To 14) As DistWageRptType
  ENumOfAct = 0
  
  'process wage distributions
  For Cnt2 = 1 To 8
    Acct$ = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    If Len(Acct$) > 0 And TransRec(1).TDist(Cnt2).DPct > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).Pct = TransRec(1).TDist(Cnt2).DPct
      EmpAct(ENumOfAct).RHrs = TransRec(1).TDist(Cnt2).DRHrs
      EmpAct(ENumOfAct).OHrs = TransRec(1).TDist(Cnt2).DOHrs
      EmpAct(ENumOfAct).RWage = TransRec(1).TDist(Cnt2).DRWage
      EmpAct(ENumOfAct).OWage = TransRec(1).TDist(Cnt2).DOWage
      EmpAct(ENumOfAct).GrossPay = OldRound#(EmpAct(ENumOfAct).RWage + EmpAct(ENumOfAct).OWage)
    End If
  Next
  'process earnings distributions
  For Cnt2 = 1 To 6
    Acct$ = QPTrim$(TransRec(1).EDist(Cnt2).EAcct)
    If Len(Acct$) > 0 Then
      ENumOfAct = ENumOfAct + 1
      LSet EmpAct(ENumOfAct).Acct = Acct$
      EmpAct(ENumOfAct).AddEarn = TransRec(1).EDist(Cnt2).EAmt
      EmpAct(ENumOfAct).GrossPay = TransRec(1).EDist(Cnt2).EAmt
    End If
  Next
  
  'HERE: got all accts for this employee
  
  'sum accts
  
  First = 1
  Start = 1
  
  Last = ENumOfAct

'purge and sum employee's dupe account distributions
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If EmpAct(First).Acct = EmpAct(Cnt2).Acct Then
        LSet EmpAct(Cnt2).Acct = ""             'blank acct num as a flag
        EmpAct(First).Pct = OldRound#(EmpAct(First).Pct + EmpAct(Cnt2).Pct)
        EmpAct(First).RHrs = OldRound#(EmpAct(First).RHrs + EmpAct(Cnt2).RHrs)
        EmpAct(First).OHrs = OldRound#(EmpAct(First).OHrs + EmpAct(Cnt2).OHrs)
        EmpAct(First).RWage = OldRound#(EmpAct(First).RWage + EmpAct(Cnt2).RWage)
        EmpAct(First).OWage = OldRound#(EmpAct(First).OWage + EmpAct(Cnt2).OWage)
        EmpAct(First).AddEarn = OldRound#(EmpAct(First).AddEarn + EmpAct(Cnt2).AddEarn)
        EmpAct(First).GrossPay = OldRound#(EmpAct(First).GrossPay + EmpAct(Cnt2).GrossPay)
      End If
    Next
Again:
    First = First + 1
  Loop Until First >= Last      '
  
  'calc percentages of matching amts to each account distribution

  For Cnt2 = 1 To ENumOfAct
    EmpAct(Cnt2).MATSocAmt = OldRound#(TransRec(1).MatchSocAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATMedAmt = OldRound#(TransRec(1).MatchMedAmt * (EmpAct(Cnt2).Pct * 0.01))
    EmpAct(Cnt2).MATRetAmt = OldRound#(TransRec(1).MatchRetAmt * (EmpAct(Cnt2).Pct * 0.01))
  Next
  '---------------------------------------------------------------------------
'calc and adjust matching distribution amts, find last active account
  'adjust Social Amt
  Do
    TotalSocAmt# = 0
    For Cnt2 = 1 To 14
      TotalSocAmt# = OldRound#(TotalSocAmt# + EmpAct(Cnt2).MATSocAmt)
    Next
    If TotalSocAmt# = 0 Then GoTo SkipSocDist
    If TotalSocAmt# <> TransRec(1).MatchSocAmt Then
      For Cnt3 = 14 To 1 Step -1
        If EmpAct(Cnt3).MATSocAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalSocAmt# > TransRec(1).MatchSocAmt Then
        DistDif# = OldRound#(TotalSocAmt# - TransRec(1).MatchSocAmt)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt - DistDif#)
      ElseIf TotalSocAmt# < TransRec(1).MatchSocAmt Then
        DistDif# = OldRound#(TransRec(1).MatchSocAmt - TotalSocAmt#)
        EmpAct(LastActive).MATSocAmt = OldRound(EmpAct(LastActive).MATSocAmt + DistDif#)
      End If
    End If
  Loop Until TotalSocAmt# = TransRec(1).MatchSocAmt

SkipSocDist:
  'adjust Medicare Amt
  Do
    TotalMedAmt# = 0
    For Cnt2 = 1 To 8
      TotalMedAmt# = OldRound#(TotalMedAmt# + EmpAct(Cnt2).MATMedAmt)
    Next

    If TotalMedAmt# = 0 Then GoTo SkipMedDist
    If TotalMedAmt# <> OldRound#(TransRec(1).MatchMedAmt) Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATMedAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalMedAmt# > TransRec(1).MatchMedAmt Then
        DistDif# = OldRound#(TotalMedAmt# - TransRec(1).MatchMedAmt)
        EmpAct(LastActive).MATMedAmt = OldRound#(EmpAct(LastActive).MATMedAmt - DistDif#)
      ElseIf TotalMedAmt# < TransRec(1).MatchMedAmt Then
        DistDif# = OldRound#(TransRec(1).MatchMedAmt - TotalMedAmt#)
        EmpAct(LastActive).MATMedAmt = OldRound#(EmpAct(LastActive).MATMedAmt + DistDif#)
      End If
    End If
  Loop Until TotalMedAmt# = OldRound#(TransRec(1).MatchMedAmt)

SkipMedDist:
  'adjust Retire Amt
  Do
    TotalRetAmt# = 0
    For Cnt2 = 1 To 8
      TotalRetAmt# = OldRound#(TotalRetAmt# + EmpAct(Cnt2).MATRetAmt)
    Next
    If TotalRetAmt# = 0 Then GoTo SkipRetDist
    If TotalRetAmt# <> TransRec(1).MatchRetAmt Then
      For Cnt3 = 8 To 1 Step -1
        If EmpAct(Cnt3).MATRetAmt > 0 Then
          LastActive = Cnt3
          Exit For
        End If
      Next
      If TotalRetAmt# > TransRec(1).MatchRetAmt Then
        DistDif# = OldRound#(TotalRetAmt# - TransRec(1).MatchRetAmt)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt - DistDif#)
      ElseIf TotalRetAmt# < TransRec(1).MatchRetAmt Then
        DistDif# = OldRound#(TransRec(1).MatchRetAmt - TotalRetAmt#)
        EmpAct(LastActive).MATRetAmt = OldRound(EmpAct(LastActive).MATRetAmt + DistDif#)
      End If
    End If
  Loop Until TotalRetAmt# = TransRec(1).MatchRetAmt

SkipRetDist:
  
'------------------------------------------------------------------------------------
'print this employee's distributions
  Print #DHandle, Dash2(1)
  DLineCnt = DLineCnt + 1
  Print #DHandle, ENumb(1) + EName(1) + BRat(1) + ORat(1)
  DLineCnt = DLineCnt + 1
  For Cnt2 = 1 To Last
    If Len(QPTrim$(EmpAct(Cnt2).Acct)) > 0 Then
      GoSub PrintEmpDist
    End If
  Next
  
  If Last > 1 Then
    GoSub PrintEmpSubTotal
  End If
'look here

  GoSub SplitDed
  
  ' sum to master distrubtion list
  For Cnt2 = 1 To Last          'process wage distributions
    Acct$ = QPTrim$(EmpAct(Cnt2).Acct)
    If Len(Acct$) > 0 Then
      If NumOfWageAccts > 0 Then
        For Cnt3 = 1 To NumOfWageAccts
          If Acct$ = QPTrim$(DistbSumAccts(Cnt3).Acct) Then
            DistbSumAccts(Cnt3).RWage = OldRound#(DistbSumAccts(Cnt3).RWage + EmpAct(Cnt2).RWage)
            DistbSumAccts(Cnt3).OWage = OldRound#(DistbSumAccts(Cnt3).OWage + EmpAct(Cnt2).OWage)
            DistbSumAccts(Cnt3).RHrs = OldRound#(DistbSumAccts(Cnt3).RHrs + EmpAct(Cnt2).RHrs)
            DistbSumAccts(Cnt3).OHrs = OldRound#(DistbSumAccts(Cnt3).OHrs + EmpAct(Cnt2).OHrs)
            DistbSumAccts(Cnt3).AddEarn = OldRound#(DistbSumAccts(Cnt3).AddEarn + EmpAct(Cnt2).AddEarn)
            DistbSumAccts(Cnt3).GrossPay = OldRound#(DistbSumAccts(Cnt3).GrossPay + EmpAct(Cnt2).GrossPay)
            DistbSumAccts(Cnt3).MATSocAmt = OldRound#(DistbSumAccts(Cnt3).MATSocAmt + EmpAct(Cnt2).MATSocAmt)
            DistbSumAccts(Cnt3).MATMedAmt = OldRound#(DistbSumAccts(Cnt3).MATMedAmt + EmpAct(Cnt2).MATMedAmt)
            DistbSumAccts(Cnt3).MATRetAmt = OldRound#(DistbSumAccts(Cnt3).MATRetAmt + EmpAct(Cnt2).MATRetAmt)
            Exit For
          End If
        Next
        If Cnt3 > NumOfWageAccts Then
          GoSub AddDistbSumAcct 'add new sum dist acct
        End If
      Else      'no previous sum accts. add new one
        GoSub AddDistbSumAcct   'add new sum dist acct
      End If
    End If
  Next
  
Return
  
AddDistbSumAcct:                'add amts to grand total acts summary
  
  NumOfWageAccts = NumOfWageAccts + 1
  If NumOfWageAccts > 1 Then
    ReDim Preserve DistbSumAccts(1 To NumOfWageAccts) As DistWageRptType
  End If
  DistbSumAccts(NumOfWageAccts) = EmpAct(Cnt2)
  
Return
  
PrintDistHeader:
  RSet Pg(1) = Str$(EPage)
  Print #DHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1)
  Print #DHandle, "Earnings Distribution"
  Print #DHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #DHandle,
  Print #DHandle, "Employee No  Name                                Base Rate   O/T Rate                          --------- Matching ----------" '+ CrLf$
  Print #DHandle, "Account Number       Sal%    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire" ' + CrLf$
  
'  Print #DHandle, Dash2(1)
  DLineCnt = DLineCnt + 7
  EPage = EPage + 1
  
Return
  
PrintEmpDist:
  LSet EDAct(1) = EmpAct(Cnt2).Acct
  RSet EDPct(1) = Using("##0.00", EmpAct(Cnt2).Pct)
  RSet EDRHrs(1) = Using(Image$, EmpAct(Cnt2).RHrs)
  RSet EDOHrs(1) = Using(Image$, EmpAct(Cnt2).OHrs)
  RSet EDRPay(1) = Using("##,##0.00", EmpAct(Cnt2).RWage)
  RSet EDOPay(1) = Using(Image$, EmpAct(Cnt2).OWage)
  RSet EDEarn(1) = Using(Image$, EmpAct(Cnt2).AddEarn)
  RSet EDGroP(1) = Using(Image$, EmpAct(Cnt2).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, EmpAct(Cnt2).MATSocAmt)
  RSet EDMAmt(1) = Using(Image$, EmpAct(Cnt2).MATMedAmt)
  RSet EDRAmt(1) = Using(Image$, EmpAct(Cnt2).MATRetAmt)
  
  Print #DHandle, EDAct(1) + EDPct(1) + EDRHrs(1) + EDOHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
  Print #DHandle,
  DLineCnt = DLineCnt + 2
Return
  
PrintEmpSubTotal:
  RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsPaid)
  RSet OTHrs(1) = Using(Image$, TransRec(1).OTHours)
  RSet EDRPay(1) = Using(Image$, TransRec(1).TotRegWage)
  RSet EDOPay(1) = Using(Image$, TransRec(1).TotOTWage)
  RSet EDEarn(1) = Using(Image$, TransRec(1).TotAdditEarn)
  RSet EDGroP(1) = Using(Image$, TransRec(1).GrossPay)
  
  RSet EDSAmt(1) = Using(Image$, TransRec(1).MatchSocAmt)
  RSet EDMAmt(1) = Using(Image$, TransRec(1).MatchMedAmt)
  RSet EDRAmt(1) = Using(Image$, TransRec(1).MatchRetAmt)
  
  'add sub totals for each employees soc, med + ret amts.
  Print #DHandle, "                   Totals" + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1) + EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
  Print #DHandle,
  DLineCnt = DLineCnt + 2
  
Return
  
PrintDistTotal:
  
  RegHrs# = 0
  TotHrs# = 0
  RegWage# = 0
  OTWage# = 0
  AddEarn# = 0
  DGPay# = 0
  
  DLineCnt = 0
  Print #DHandle, FF$
  
  If NumOfWageAccts > 1 Then
    SortT DistbSumAccts(), NumOfWageAccts
  End If
  
  'sort distribution account numbers (page 5)
  'print the Summary of ALL Distrubtion Accounts
  GoSub PrintAcctSumHeader
  
  'ON LOCAL ERROR RESUME NEXT
  
  For cnt = 1 To NumOfWageAccts
    LSet EDAct(1) = DistbSumAccts(cnt).Acct
    If GPay# <= 0 Then Return
    RSet EDPct(1) = Using(Image4$, (DistbSumAccts(cnt).GrossPay / GPay#) * 100)
    RSet RHrs(1) = Using(Image3$, DistbSumAccts(cnt).RHrs)
    RegHrs# = OldRound#(RegHrs# + DistbSumAccts(cnt).RHrs)
    RSet OTHrs(1) = Using(Image3$, DistbSumAccts(cnt).OHrs)
    TotHrs# = OldRound#(TotHrs# + DistbSumAccts(cnt).OHrs)
    RSet EDRPay(1) = Using(Image3$, DistbSumAccts(cnt).RWage)
    RegWage# = OldRound#(RegWage# + DistbSumAccts(cnt).RWage)
    RSet EDOPay(1) = Using(Image3$, DistbSumAccts(cnt).OWage)
    OTWage# = OldRound#(OTWage# + DistbSumAccts(cnt).OWage)
    RSet EDEarn(1) = Using(Image3$, DistbSumAccts(cnt).AddEarn)
    AddEarn# = OldRound#(AddEarn# + DistbSumAccts(cnt).AddEarn)
    RSet EDGroP(1) = Using(Image3$, DistbSumAccts(cnt).GrossPay)
    DGPay# = OldRound#(DGPay# + DistbSumAccts(cnt).GrossPay)
    
    RSet EDSAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATSocAmt)
    ASAmt# = OldRound#(ASAmt# + DistbSumAccts(cnt).MATSocAmt)
    
    RSet EDMAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATMedAmt)
    AMAmt# = OldRound#(AMAmt# + DistbSumAccts(cnt).MATMedAmt)
    
    RSet EDRAmt(1) = Using(Image3$, DistbSumAccts(cnt).MATRetAmt)
    ARAmt# = OldRound#(ARAmt# + DistbSumAccts(cnt).MATRetAmt)
    
    Print #DHandle, EDAct(1) + EDPct(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1),
    Print #DHandle, EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1)
    
    DLineCnt = DLineCnt + 1
'    If DLineCnt >= DMaxLines Then '8/15
'      DLineCnt = 0
'      GoSub PrintAcctSumHeader
'    End If
  Next
  'ON LOCAL ERROR GOTO 0
  
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet OTHrs(1) = Using(Image3$, TotHrs#)
  RSet EDRPay(1) = Using(Image3$, RegWage#)
  RSet EDOPay(1) = Using(Image3$, OTWage#)
  RSet EDEarn(1) = Using(Image3$, AddEarn#)
  RSet EDGroP(1) = Using(Image3$, DGPay#)
  
  RSet EDSAmt(1) = Using(Image3$, ASAmt#)
  RSet EDMAmt(1) = Using(Image3$, AMAmt#)
  RSet EDRAmt(1) = Using(Image3$, ARAmt#)
  
  Print #DHandle, Dash2(1)

  Fill11(1) = ""
  Print #DHandle, "Totals        " + Fill11(1) + RHrs(1) + OTHrs(1) + EDRPay(1) + EDOPay(1),
  Print #DHandle, EDEarn(1) + EDGroP(1) + EDSAmt(1) + EDMAmt(1) + EDRAmt(1) ' + CrLf$
  Print #DHandle, FF$
  
Return
  
PrintAcctSumHeader:
  Print #DHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1)
  Print #DHandle, "Earnings Distribution Totals"
  Print #DHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  Print #DHandle, "                                                                                               --------- Matching ----------"
  Print #DHandle, "Account Number        Pct    Reg Hrs    O/T Hrs    Reg Pay    O/T Pay  Tot Other  Gross Pay    Soc Sec   Medicare     Retire"
  Print #DHandle, Dash2(1)
  DLineCnt = DLineCnt + 6
Return
  
PrintSplitTotal:

  If SplDedCnt = 0 Then SplDedCnt = 1

  ReDim Preserve SplitDeds(1 To SplDedCnt) As SplitDedsType

  If SplDedCnt > 1 Then
    SortE SplitDeds(), SplDedCnt
  End If
    

  EPage = EPage + 1
  RSet Pg(1) = Str$(EPage)

  Print #DHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1)
  Print #DHandle, "Fund Distribution Totals"
  Print #DHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
  
  Print #DHandle, "Fund           Fed      Sta     Med       Soc      Ret"
  Print #DHandle, "Deductions       1 thru " & NumOfDeds
  
  Print #DHandle, Dash2(1)
  
'look here
  For CCnt3 = 1 To SplDedCnt
    If Len(QPTrim$(SplitDeds(CCnt3).Acct)) = 0 Then GoTo SkipSplLine
    '5/1 This is the total by fund
    LSet SplDed2(1) = SplitDeds(CCnt3).Acct
    TempLine$(CCnt3) = SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).FEDTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).STATAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).MEDTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).SOCTAX)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).RetTax)
    TempLine$(CCnt3) = TempLine$(CCnt3) + SplDed2(1)
    Print #DHandle, TempLine$(CCnt3) '+ CrLf$
    TempLine$(CCnt3) = ""
    
    For x = 1 To 5
      TempLine$(x) = ""
    Next x
    Nextx = 1
    tripCnt = 1
    For DCnt = 1 To NumOfDeds
      If tripCnt = 13 Then
        tripCnt = 1
        Nextx = Nextx + 1
      End If
      RSet SplDed2(1) = Using("###0.00", SplitDeds(CCnt3).DedAmts(DCnt))
      TempLine$(Nextx) = TempLine$(Nextx) + SplDed2(1)
      tripCnt = tripCnt + 1
    Next
    
    For x = 1 To Nextx
      Print #DHandle, TempLine$(x)
    Next x
    
SkipSplLine:
  Next
  Print #DHandle, Dash2(1)
  Print #DHandle, FF$
Return
  
SplitDed:
  EmpSplCnt = 0
  If ENumOfAct = 0 Then ENumOfAct = 1
  ReDim EmpSpl(1 To ENumOfAct) As SplitDedsType
  
  For CCnt3 = 1 To ENumOfAct
    SFund$ = Left$(EmpAct(CCnt3).Acct, 2)
    If SFund$ = "  " Then
    Else
      GoSub CheckEmpSpl
      GoSub CheckFund
    End If
  Next
  
  GoSub FixSplitBal
  
  '----------------------------New Void Check Code--6/15/2004-----------

  ReDim ThisSplAcct(1 To EmpSplCnt) As String
  
  NextSpl = 1
  For x = 1 To EmpSplCnt
    ThisSplAcct(x) = EmpSpl(x).Acct
  Next x
  
  TempVoid.EmpNum = QPTrim$(EmpRec1(1).EmpNo)
  TempVoid.NumOfAccts = 0
  
  For CCnt3 = 1 To ENumOfAct
    If Len(QPTrim$(EmpAct(CCnt3).Acct)) <> 0 Then
      TempVoid.NumOfAccts = TempVoid.NumOfAccts + 1
    End If
  Next CCnt3
  
  For CCnt3 = 1 To ENumOfAct 'EmpSplCnt
    ThisPR = EmpAct(CCnt3).GrossPay
    If Len(QPTrim$(EmpAct(CCnt3).Acct)) = 0 Then GoTo NoAcctNum
    EmpAct(CCnt3).Acct = ReplaceString(EmpAct(CCnt3).Acct, "-", "")
    FundNumOnly$ = QPTrim(Mid(EmpAct(CCnt3).Acct, 1, FundLen))
    FundAndAcctOnly$ = QPTrim(Mid(EmpAct(CCnt3).Acct, 1, FACnt))

    TempVoid.SOCMATDBAmt = EmpAct(CCnt3).MATSocAmt
    DSocMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).SOCEXP)
    DSocMatchAcct$ = ReplaceString(DSocMatchAcct$, "-", "")
    DSocMatchAcct$ = AddDashesToGLNumber(DSocMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.SOCMATDBGL = DSocMatchAcct$

    TempVoid.MEDMATDBAmt = EmpAct(CCnt3).MATMedAmt
    DMedMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).MEDEXP)
    DMedMatchAcct$ = ReplaceString(DMedMatchAcct$, "-", "")
    DMedMatchAcct$ = AddDashesToGLNumber(DMedMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.MEDMATDBGL = DMedMatchAcct$

    TempVoid.RETMATDBAmt = EmpAct(CCnt3).MATRetAmt
    DRetMatchAcct$ = FundAndAcctOnly$ + QPTrim$(SysRec(1).RETEXP)
    DRetMatchAcct$ = ReplaceString(DRetMatchAcct$, "-", "")
    DRetMatchAcct$ = AddDashesToGLNumber(DRetMatchAcct$, FundLen, AcctLen, DetLen)
    TempVoid.RETMATDBGL = DRetMatchAcct$

    DWagesAcct$ = QPTrim$(EmpAct(CCnt3).Acct)
    DWagesAcct$ = ReplaceString(DWagesAcct$, "-", "")
    DWagesAcct$ = AddDashesToGLNumber(DWagesAcct$, FundLen, AcctLen, DetLen)
    TempVoid.WagesGL = DWagesAcct$
    TempVoid.WagesAmt = EmpAct(CCnt3).GrossPay

    For NextSpl = 1 To EmpSplCnt
      If QPTrim$(ThisSplAcct(NextSpl)) = Mid(EmpAct(CCnt3).Acct, 1, FundLen) Then
        ThisSplAcct(NextSpl) = "" 'once read it cannot be read again
      TempVoid.FEDWHAmt = EmpSpl(CCnt3).FEDTAX
      ThisPR = ThisPR - EmpSpl(CCnt3).FEDTAX
      CFedWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(1).Acct)
      CFedWHAcct$ = ReplaceString(CFedWHAcct$, "-", "")
      CFedWHAcct$ = AddDashesToGLNumber(CFedWHAcct$, FundLen, AcctLen, DetLen)
      TempVoid.FEDWHGL = CFedWHAcct$
  
      TempVoid.STAWHAmt = EmpSpl(CCnt3).STATAX
      ThisPR = ThisPR - EmpSpl(CCnt3).STATAX
      CStaWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(2).Acct)
      CStaWHAcct$ = ReplaceString(CStaWHAcct$, "-", "")
      CStaWHAcct$ = AddDashesToGLNumber(CStaWHAcct$, FundLen, AcctLen, DetLen)
      TempVoid.STAWHGL = CStaWHAcct$
  
      TempVoid.MEDWHAmt = EmpSpl(CCnt3).MEDTAX
      CMedMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).MEDLIAB)
      ThisPR = ThisPR - EmpSpl(CCnt3).MEDTAX
      CMedMatchAcct$ = ReplaceString(CMedMatchAcct$, "-", "")
      CMedMatchAcct$ = AddDashesToGLNumber(CMedMatchAcct$, FundLen, AcctLen, DetLen)
      TempVoid.MEDMATCRGL = CMedMatchAcct$
      TempVoid.MEDMATCRAmt = EmpAct(CCnt3).MATMedAmt
      CMedWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(4).Acct)
      CMedWHAcct$ = ReplaceString(CMedWHAcct$, "-", "")
      CMedWHAcct$ = AddDashesToGLNumber(CMedWHAcct$, FundLen, AcctLen, DetLen)
      TempVoid.MEDWHGL = CMedWHAcct$
  
      TempVoid.SOCWHAmt = EmpSpl(CCnt3).SOCTAX
      ThisPR = ThisPR - EmpSpl(CCnt3).SOCTAX
      CSocMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).SOCLIAB)
      CSocMatchAcct$ = ReplaceString(CSocMatchAcct$, "-", "")
      CSocMatchAcct$ = AddDashesToGLNumber(CSocMatchAcct$, FundLen, AcctLen, DetLen)
      TempVoid.SOCMATCRGL = CSocMatchAcct$
      TempVoid.SOCMATCRAmt = EmpAct(CCnt3).MATSocAmt
      CSocWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(3).Acct)
      CSocWHAcct$ = ReplaceString(CSocWHAcct$, "-", "")
      CSocWHAcct$ = AddDashesToGLNumber(CSocWHAcct$, FundLen, AcctLen, DetLen)
      TempVoid.SOCWHGL = CSocWHAcct$
  
      TempVoid.RETWHAmt = EmpSpl(CCnt3).RetTax
      ThisPR = ThisPR - EmpSpl(CCnt3).RetTax
      CRetMatchAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).RETLIAB)
      CRetMatchAcct$ = ReplaceString(CRetMatchAcct$, "-", "")
      CRetMatchAcct$ = AddDashesToGLNumber(CRetMatchAcct$, FundLen, AcctLen, DetLen)
      TempVoid.RETMATCRGL = CRetMatchAcct$
      TempVoid.RETMATCRAmt = EmpAct(CCnt3).MATRetAmt
      CRetWHAcct$ = FundNumOnly$ + QPTrim$(SysRec(1).Liab(5).Acct)
      CRetWHAcct$ = ReplaceString(CRetWHAcct$, "-", "")
      CRetWHAcct$ = AddDashesToGLNumber(CRetWHAcct$, FundLen, AcctLen, DetLen)
      TempVoid.RETWHGL = CRetWHAcct$
  
      For x = 1 To 50
        If EmpSpl(CCnt3).DedAmts(x) > 0 Then
          CDedAcct$ = FundNumOnly$ + DedCodeNums(x)
          CDedAcct$ = ReplaceString(CDedAcct$, "-", "")
          CDedAcct$ = AddDashesToGLNumber(CDedAcct$, FundLen, AcctLen, DetLen)
          TempVoid.DedData(x).DAmt = EmpSpl(CCnt3).DedAmts(x)
          ThisPR = ThisPR - EmpSpl(CCnt3).DedAmts(x)
          TempVoid.DedData(x).DedGLNum = CDedAcct$
          TempVoid.DedData(x).DedDesc = "VP" + QPTrim$(DedDescTrimmed(x))
        Else
          TempVoid.DedData(x).DAmt = 0
          TempVoid.DedData(x).DedGLNum = ""
        End If
      Next x
        Exit For
      End If
    Next NextSpl
    
    If NextSpl = EmpSplCnt + 1 Then
      TempVoid.FEDWHAmt = 0
      TempVoid.STAWHAmt = 0
      TempVoid.MEDWHAmt = 0
      TempVoid.SOCWHAmt = 0
      TempVoid.RETWHAmt = 0
      For x = 1 To 50
        TempVoid.DedData(x).DAmt = 0
      Next x
    End If
    
    TempVoid.VoidFlag = False
    TempVoid.CheckAmt = 0
    TempVoid.CheckDate = 0
    TempVoid.CheckNum = "0"
    TempVoid.TransRec = 0
    TempVoid.PRNet = ThisPR
    TempVoid.PRNetGL = FundNumOnly + QPTrim$(SysRec(1).CashAcct)
    TempVoid.PRNetGL = ReplaceString(TempVoid.PRNetGL, "-", "")
    TempVoid.PRNetGL = AddDashesToGLNumber(TempVoid.PRNetGL, FundLen, AcctLen, DetLen)
    TempVoid.Type = QPTrim$(SysRec(1).USEIMP)
    TempVoid.Pad = ""
    If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
      TempVoid.PPEAmt = ThisPR
      TempVoid.PPEGL = QPTrim$(SysRec(1).ICRACCT) + FundNumOnly
      TempVoid.PPEGL = ReplaceString(TempVoid.PPEGL, "-", "")
      TempVoid.PPEGL = AddDashesToGLNumber(TempVoid.PPEGL, FundLen, AcctLen, DetLen)
      TempVoid.PPETotAmt = ThisPR
      TempVoid.PPETotGL = QPTrim$(SysRec(1).IDRACCT)
      TempVoid.PPETotGL = ReplaceString(TempVoid.PPETotGL, "-", "")
      TempVoid.PPETotGL = AddDashesToGLNumber(TempVoid.PPETotGL, FundLen, AcctLen, DetLen)
    Else
      TempVoid.PPEAmt = 0
      TempVoid.PPEGL = ""
      TempVoid.PPETotAmt = 0
      TempVoid.PPETotGL = ""
    End If
    TVCnt = TVCnt + 1
    TempVoid.EmpNum = TempVoid.EmpNum
    Put TVHandle, TVCnt, TempVoid
NoAcctNum:

  Next CCnt3
  
  '----------------^^^^--------New Void Check Code--6/15/2004-----------
  Print #DHandle, "Deductions by Fund"
  AddLine = 0
  For x = 1 To 5
    If Len(QPTrim$(STitle$(x))) > 0 Then
      Print #DHandle, STitle$(x)
      AddLine = AddLine + 1
    End If
  Next x
  
'  DLineCnt = DLineCnt + 1 + AddLine
  DLineCnt = DLineCnt + AddLine '+ 1
  
  '******
  ReDim EmpDedSum(1 To (5 + NumOfDeds)) As Double '17) As Double changed 5/1
  tripCnt = 1
  Nextx = 1
  For CCnt3 = 1 To EmpSplCnt
    For x = 1 To 5
      TempLine$(x) = ""
    Next x
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    LSet SplDed(1) = EmpSpl(CCnt3).Acct
    TempLine$(Nextx) = Left$(SplDed(1), 5)
    For DCnt = 1 To NumOfDeds
      If tripCnt = 13 Then GoTo EndLine1
      RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).DedAmts(DCnt))
      EmpDedSum(DCnt) = OldRound(EmpDedSum(DCnt) + EmpSpl(CCnt3).DedAmts(DCnt))
      TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
      tripCnt = tripCnt + 1
    Next
EndLine1:
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).FEDTAX)
    EmpDedSum(NumOfDeds + 1) = OldRound(EmpDedSum(NumOfDeds + 1) + EmpSpl(CCnt3).FEDTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).STATAX)
    EmpDedSum(NumOfDeds + 2) = OldRound(EmpDedSum(NumOfDeds + 2) + EmpSpl(CCnt3).STATAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).MEDTAX)
    EmpDedSum(NumOfDeds + 3) = OldRound(EmpDedSum(NumOfDeds + 3) + EmpSpl(CCnt3).MEDTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).SOCTAX)
    EmpDedSum(NumOfDeds + 4) = OldRound(EmpDedSum(NumOfDeds + 4) + EmpSpl(CCnt3).SOCTAX)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).RetTax)
    EmpDedSum(NumOfDeds + 5) = OldRound(EmpDedSum(NumOfDeds + 5) + EmpSpl(CCnt3).RetTax)
    TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    
    Print #DHandle, TempLine$(Nextx)
    DLineCnt = DLineCnt + 1
    tripCnt = tripCnt + 1
  
    If tripCnt > 12 Then
      tripCnt = 1
      Nextx = Nextx + 1
      For x = 13 To NumOfDeds
        If tripCnt = 13 Then
          tripCnt = 1
          Nextx = Nextx + 1
        End If
        RSet SplDed(1) = Using("##0.00", EmpSpl(CCnt3).DedAmts(x))
        EmpDedSum(x) = OldRound(EmpDedSum(x) + EmpSpl(CCnt3).DedAmts(x))
        If tripCnt = 1 Then
          TempLine$(Nextx) = TempLine$(Nextx) + "     " + SplDed(1)
        Else
          TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
        End If
        tripCnt = tripCnt + 1
      Next x  'Start Here 10/7
      
      For x = 2 To 5 'added 5/28 ...had For x = Nextx to 5 but
      'discovered that line 1 always prints with the taxes
      'on the end so this section should always begin a 2
      'and if there are no values on line 2 then this should
      'handle it
        If Len(QPTrim$(TempLine$(x))) > 0 Then
          Print #DHandle, TempLine(x)
          DLineCnt = DLineCnt + 1 ' 8/15
       End If
      Next x
    End If
    tripCnt = 1
    Nextx = 1 'added 10/10
  Next
  
  For x = 1 To 5
    TempLine(x) = ""
  Next x
  
  Print #DHandle, "Total"
  DLineCnt = DLineCnt + 1 '8/15
  tripCnt = 1
  Nextx = 1
  For x = 1 To (5 + NumOfDeds) '17
    RSet SplDed(1) = Using("##0.00", EmpDedSum(x))
    If tripCnt = 1 Then
      TempLine$(Nextx) = TempLine$(Nextx) + "     " + SplDed(1)
    Else
      TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
    End If
    tripCnt = tripCnt + 1
    If tripCnt = 13 Then Exit For
  Next
  
  For x = (NumOfDeds + 1) To (NumOfDeds + 5)
    RSet SplDed(1) = Using("##0.00", EmpDedSum(x))
    TempLine(Nextx) = TempLine(Nextx) + SplDed(1)
  Next x
  Print #DHandle, TempLine(Nextx)
  DLineCnt = DLineCnt + 1 '8/15
  
  If tripCnt > 12 Then
    tripCnt = 1
    Nextx = 2
    For x = 13 To NumOfDeds
      If tripCnt = 13 Then
        tripCnt = 1
        Nextx = Nextx + 1
      End If
      RSet SplDed(1) = Using("##0.00", EmpDedSum(x))
      If tripCnt = 1 Then
        TempLine$(Nextx) = TempLine$(Nextx) + "     " + SplDed(1)
      Else
        TempLine$(Nextx) = TempLine$(Nextx) + SplDed(1)
      End If
      tripCnt = tripCnt + 1
    Next x
    
    For x = 2 To 5
      If Len(QPTrim$(TempLine$(x))) > 0 Then
        Print #DHandle, TempLine$(x)
        DLineCnt = DLineCnt + 1 '8/15
      End If
    Next x
  End If
  
  If DLineCnt > DMaxLines Then '8/15
    DLineCnt = 0
    Print #DHandle, FF$
    GoSub PrintDistHeader
  End If
    
Return
  
CheckEmpSpl:

  If EmpSplCnt = 0 Then
    GoSub AddEmpSpl
  End If
  GotIt = False
  
'look here
'6/10/03...uncovered a rounding error with some split combinations (if there
'was a 50/50 split with 2 accounts and the state tax was also $50 there
'was an error). Added OldRound below which further refined the numbers and
'fixed rounding errors
  For ESCnt = 1 To EmpSplCnt
    If Left$(EmpSpl(ESCnt).Acct, 2) = SFund$ Then
      GotIt = True
      SPct@ = OldRound(EmpAct(CCnt3).Pct * 0.01) '6/10/03 added OldRound
      If SPct@ > 0 Then
        For DCnt = 1 To NumOfDeds
          ESplAmt# = OldRound(TransRec(1).DAmt(DCnt) * SPct@)
          EmpSpl(ESCnt).DedAmts(DCnt) = OldRound(EmpSpl(ESCnt).DedAmts(DCnt) + ESplAmt#)
        Next
        'look here
'        EmpSpl(ESCnt).FEDTAX = OldRound(EmpSpl(ESCnt).FEDTAX + (OldRound(TransRec(1).FedTaxAmt) * SPct@)) '6/10/03
        EmpSpl(ESCnt).FEDTAX = OldRound(EmpSpl(ESCnt).FEDTAX + OldRound(TransRec(1).FedTaxAmt * SPct@))
'        EmpSpl(ESCnt).STATAX = OldRound(EmpSpl(ESCnt).STATAX + (TransRec(1).StaTaxAmt) * SPct@)
        EmpSpl(ESCnt).STATAX = OldRound(EmpSpl(ESCnt).STATAX + OldRound(TransRec(1).StaTaxAmt * SPct@)) '6/10/03
'        EmpSpl(ESCnt).MEDTAX = OldRound(EmpSpl(ESCnt).MEDTAX + (TransRec(1).MedTaxAmt) * SPct@) '6/10/03
        EmpSpl(ESCnt).MEDTAX = OldRound(EmpSpl(ESCnt).MEDTAX + OldRound(TransRec(1).MedTaxAmt * SPct@))
'        EmpSpl(ESCnt).SOCTAX = OldRound(EmpSpl(ESCnt).SOCTAX + (TransRec(1).SocTaxAmt) * SPct@)
        EmpSpl(ESCnt).SOCTAX = OldRound(EmpSpl(ESCnt).SOCTAX + OldRound(TransRec(1).SocTaxAmt * SPct@)) '6/10/03
'        EmpSpl(ESCnt).RetTax = OldRound(EmpSpl(ESCnt).RetTax + (TransRec(1).RetireAmt) * SPct@)
        EmpSpl(ESCnt).RetTax = OldRound(EmpSpl(ESCnt).RetTax + OldRound(TransRec(1).RetireAmt * SPct@)) '6/10/03
'        EmpSpl(ESCnt).EICPay = OldRound(EmpSpl(ESCnt).EICPay + (TransRec(1).EICAmt) * SPct@)
        EmpSpl(ESCnt).EICPay = OldRound(EmpSpl(ESCnt).EICPay + OldRound(TransRec(1).EICAmt * SPct@)) '6/10/03
      End If
    End If
  Next
  
  If Not GotIt Then
    GoSub AddEmpSpl
    GoTo CheckEmpSpl
  Else
    GotIt = False
  End If

Return
 
AddEmpSpl:
  EmpSplCnt = EmpSplCnt + 1
  EmpSpl(EmpSplCnt).Acct = SFund$
Return
  
CheckFund:
  If SplDedCnt = 0 Then
    GoSub AddNewFund
  End If
  
  GotIt = False
  Nextx = 1
  tripCnt = 1
'6/10/03...uncovered a rounding error with some split combinations (if there
'was a 50/50 split with 2 accounts and the state tax was also $50 there
'was an error). Added OldRound below which further refined the numbers and
'fixed rounding errors
  For CCnt01 = 1 To SplDedCnt
    If Left$(SplitDeds(CCnt01).Acct, 2) = SFund$ Then
      GotIt = True
      SPct@ = OldRound(EmpAct(CCnt3).Pct * 0.01) '6/10/03 added OldRound
      LSet EDAct(1) = SFund$
      TempLine$(Nextx) = ""
      
      For DCnt = 1 To NumOfDeds
        SDedAmt# = OldRound(TransRec(1).DAmt(DCnt) * SPct@)
        SplitDeds(CCnt01).DedAmts(DCnt) = OldRound(SplitDeds(CCnt01).DedAmts(DCnt) + SDedAmt#)
      Next
      
'      SplitDeds(CCnt01).FEDTAX = OldRound(SplitDeds(CCnt01).FEDTAX + OldRound(TransRec(1).FedTaxAmt) * SPct@) '6/10/03
      SplitDeds(CCnt01).FEDTAX = OldRound(SplitDeds(CCnt01).FEDTAX + OldRound(TransRec(1).FedTaxAmt * SPct@))
'      SplitDeds(CCnt01).STATAX = OldRound(SplitDeds(CCnt01).STATAX + (TransRec(1).StaTaxAmt) * SPct@)
      SplitDeds(CCnt01).STATAX = OldRound(SplitDeds(CCnt01).STATAX + OldRound(TransRec(1).StaTaxAmt * SPct@)) '6/10/03
'      SplitDeds(CCnt01).MEDTAX = OldRound(SplitDeds(CCnt01).MEDTAX + (TransRec(1).MedTaxAmt) * SPct@) '6/10/03
      SplitDeds(CCnt01).MEDTAX = OldRound(SplitDeds(CCnt01).MEDTAX + OldRound(TransRec(1).MedTaxAmt * SPct@))
'      SplitDeds(CCnt01).SOCTAX = OldRound(SplitDeds(CCnt01).SOCTAX + (TransRec(1).SocTaxAmt) * SPct@)
      SplitDeds(CCnt01).SOCTAX = OldRound(SplitDeds(CCnt01).SOCTAX + OldRound(TransRec(1).SocTaxAmt * SPct@)) '6/10/03
'      SplitDeds(CCnt01).RetTax = OldRound(SplitDeds(CCnt01).RetTax + (TransRec(1).RetireAmt) * SPct@)
      SplitDeds(CCnt01).RetTax = OldRound(SplitDeds(CCnt01).RetTax + OldRound(TransRec(1).RetireAmt * SPct@)) '6/10/03
'      SplitDeds(CCnt01).EICPay = OldRound(SplitDeds(CCnt01).EICPay + (TransRec(1).EICAmt) * SPct@)
      SplitDeds(CCnt01).EICPay = OldRound(SplitDeds(CCnt01).EICPay + OldRound(TransRec(1).EICAmt * SPct@)) '6/10/03
      Exit For
    End If
  Next

  If Not GotIt Then
    GoSub AddNewFund
    GoTo CheckFund
  Else
    GotIt = False
  End If
  
Return
AddNewFund:
  SplDedCnt = SplDedCnt + 1
  If SplDedCnt > ArraySize Then
    ArraySize = ArraySize + 5
    ReDim Preserve SplitDeds(1 To ArraySize) As SplitDedsType
  End If
  
  SplitDeds(SplDedCnt).Acct = SFund$
Return
  
FixSplitBal:

  tFedTax# = 0
  tStaTax# = 0
  tMedTax# = 0
  tSocTax# = 0
  tRetTax# = 0
  tEICPay# = 0
  
  For CCnt = 1 To NumOfDeds
    DedAmt# = 0
    For aa = 1 To EmpSplCnt
      DedAmt# = OldRound(DedAmt# + EmpSpl(aa).DedAmts(CCnt))
    Next
    If DedAmt# > TransRec(1).DAmt(CCnt) Then
      DistDif# = OldRound(DedAmt# - TransRec(1).DAmt(CCnt))
      EmpSpl(EmpSplCnt).DedAmts(CCnt) = OldRound(EmpSpl(EmpSplCnt).DedAmts(CCnt) - DistDif#)
      SplitDeds(EmpSplCnt).DedAmts(CCnt) = OldRound(SplitDeds(EmpSplCnt).DedAmts(CCnt) - DistDif#)
    ElseIf DedAmt# < TransRec(1).DAmt(CCnt) Then
      DistDif# = OldRound(TransRec(1).DAmt(CCnt) - DedAmt#)
      EmpSpl(EmpSplCnt).DedAmts(CCnt) = OldRound(EmpSpl(EmpSplCnt).DedAmts(CCnt) + DistDif#)
      SplitDeds(EmpSplCnt).DedAmts(CCnt) = OldRound(SplitDeds(EmpSplCnt).DedAmts(CCnt) + DistDif#)
    End If
  Next
  
  For aa = 1 To EmpSplCnt
    tFedTax# = OldRound(tFedTax# + EmpSpl(aa).FEDTAX)
    tStaTax# = OldRound(tStaTax# + EmpSpl(aa).STATAX)
    tMedTax# = OldRound(tMedTax# + EmpSpl(aa).MEDTAX)
    tSocTax# = OldRound(tSocTax# + EmpSpl(aa).SOCTAX)
    tRetTax# = OldRound(tRetTax# + EmpSpl(aa).RetTax)
    tEICPay# = OldRound(tEICPay# + EmpSpl(aa).EICPay)
  Next
  
  If tFedTax# > TransRec(1).FedTaxAmt Then
    'STOP
    DistDif# = OldRound(tFedTax# - TransRec(1).FedTaxAmt)
    EmpSpl(EmpSplCnt).FEDTAX = OldRound(EmpSpl(EmpSplCnt).FEDTAX - DistDif#)
    SplitDeds(EmpSplCnt).FEDTAX = OldRound(SplitDeds(EmpSplCnt).FEDTAX - DistDif#)
  ElseIf tFedTax# < TransRec(1).FedTaxAmt Then
    'STOP
    DistDif# = OldRound(TransRec(1).FedTaxAmt - tFedTax#)
    EmpSpl(EmpSplCnt).FEDTAX = OldRound(EmpSpl(EmpSplCnt).FEDTAX + DistDif#)
    SplitDeds(EmpSplCnt).FEDTAX = OldRound#(SplitDeds(EmpSplCnt).FEDTAX + DistDif#)
  End If

  If tStaTax# > TransRec(1).StaTaxAmt Then
    DistDif# = OldRound(tStaTax# - TransRec(1).StaTaxAmt)
    EmpSpl(EmpSplCnt).STATAX = OldRound(EmpSpl(EmpSplCnt).STATAX - DistDif#)
    SplitDeds(EmpSplCnt).STATAX = OldRound(SplitDeds(EmpSplCnt).STATAX - DistDif#)
  ElseIf tStaTax# < TransRec(1).StaTaxAmt Then
    DistDif# = OldRound(TransRec(1).StaTaxAmt - tStaTax#)
    EmpSpl(EmpSplCnt).STATAX = OldRound(EmpSpl(EmpSplCnt).STATAX + DistDif#)
    SplitDeds(EmpSplCnt).STATAX = OldRound(SplitDeds(EmpSplCnt).STATAX + DistDif#)
  End If
  
  If tMedTax# > TransRec(1).MedTaxAmt Then
    DistDif# = OldRound(tMedTax# - TransRec(1).MedTaxAmt)
    EmpSpl(EmpSplCnt).MEDTAX = OldRound(EmpSpl(EmpSplCnt).MEDTAX - DistDif#)
    SplitDeds(EmpSplCnt).MEDTAX = OldRound(SplitDeds(EmpSplCnt).MEDTAX - DistDif#)
  ElseIf tMedTax# < TransRec(1).MedTaxAmt Then
    DistDif# = OldRound(TransRec(1).MedTaxAmt - tMedTax#)
    EmpSpl(EmpSplCnt).MEDTAX = OldRound(EmpSpl(EmpSplCnt).MEDTAX + DistDif#)
    SplitDeds(EmpSplCnt).MEDTAX = OldRound(SplitDeds(EmpSplCnt).MEDTAX + DistDif#)
  End If
  
  If tSocTax# > TransRec(1).SocTaxAmt Then
    DistDif# = OldRound(tSocTax# - TransRec(1).SocTaxAmt)
    EmpSpl(EmpSplCnt).SOCTAX = OldRound(EmpSpl(EmpSplCnt).SOCTAX - DistDif#)
    SplitDeds(EmpSplCnt).SOCTAX = OldRound(SplitDeds(EmpSplCnt).SOCTAX - DistDif#)
  ElseIf tSocTax# < TransRec(1).SocTaxAmt Then
    DistDif# = OldRound(TransRec(1).SocTaxAmt - tSocTax#)
    EmpSpl(EmpSplCnt).SOCTAX = OldRound(EmpSpl(EmpSplCnt).SOCTAX + DistDif#)
    SplitDeds(EmpSplCnt).SOCTAX = OldRound(SplitDeds(EmpSplCnt).SOCTAX + DistDif#)
  End If
  
  If tRetTax# > TransRec(1).RetireAmt Then
    DistDif# = OldRound(tRetTax# - TransRec(1).RetireAmt)
    EmpSpl(EmpSplCnt).RetTax = OldRound(EmpSpl(EmpSplCnt).RetTax - DistDif#)
    SplitDeds(EmpSplCnt).RetTax = OldRound(SplitDeds(EmpSplCnt).RetTax - DistDif#)

  ElseIf tRetTax# < TransRec(1).RetireAmt Then
    DistDif# = OldRound(TransRec(1).RetireAmt - tRetTax#)
    EmpSpl(EmpSplCnt).RetTax = OldRound(EmpSpl(EmpSplCnt).RetTax + DistDif#)
    SplitDeds(EmpSplCnt).RetTax = OldRound(SplitDeds(EmpSplCnt).RetTax + DistDif#)
  End If
  
  If tEICPay# > TransRec(1).EICAmt Then
    DistDif# = OldRound(tEICPay# - TransRec(1).EICAmt)
    EmpSpl(EmpSplCnt).EICPay = OldRound(EmpSpl(EmpSplCnt).EICPay - DistDif#)
    SplitDeds(EmpSplCnt).EICPay = OldRound(SplitDeds(EmpSplCnt).EICPay - DistDif#)
  ElseIf tEICPay# < TransRec(1).EICAmt Then
    DistDif# = OldRound(TransRec(1).EICAmt - tEICPay#)
    EmpSpl(EmpSplCnt).EICPay = OldRound(EmpSpl(EmpSplCnt).EICPay + DistDif#)
    SplitDeds(EmpSplCnt).EICPay = OldRound(SplitDeds(EmpSplCnt).EICPay + DistDif#)
  End If
  
Return
  
End Sub

Sub MakeGLIFFileST(SplitDeds() As SplitDedsType, DistbSumAccts() As DistWageRptType, TotEIC#, TotDeds#(), Passed#(), PrintGLRpt)

  Dim FEDWith$, MEDWith$, MEDMatLib$, SOCWith$
  Dim SOCMatLib$, RETWith$, RETMatLib$, STAWith$
  Dim MEDMat$, SOCMat$, RETMat$
  Dim Image3$, FACnt As Integer
  Dim NumOfWageAccts As Integer
  Dim SplDedCnt As Integer, GLIFTDate$
  Dim FedAcct$, StaAcct$, SocAcct$, MedAcct$
  Dim RetAcct$, GLIFSource$, cnt As Integer
  Dim GLIFArrayAdj As Integer, NextAcct As Integer
  Dim SysCash$, ZZCnt As Integer
  Dim TAcct$, NumDFunds As Integer
  Dim CurrFund$, FirstC As Integer
  Dim LastC As Integer, NumCFunds As Integer
  Dim TotalFunds As Integer, ThisFund As Integer
  Dim First As Integer, Start As Integer, Last As Integer
  Dim Cnt2 As Integer, FundCash$
  Dim TotalGLIFS As Integer, THandle As Integer
  Dim AcctNum$, SOCEXP$, MEDEXP$, RETEXP$
  Dim SOCLIAB$, MEDLIAB$, RETLIAB$
  Dim GLIFRecLen As Integer, RecNo&
  Dim GHandle As Integer, RptName$
  Dim TotDebit#
  Dim GLIdxRecLen As Integer, TotCredit#
  Dim OutOfOrder As Boolean, AcctOk As Boolean
  Dim ActualAccts As Integer, NoAcctNum As Integer
  Dim ArraySize As Integer, GLIdxName$
  Dim Lines As Integer
  Dim Fund$, FDebit#, FCredit#, NFund$, FF$
  Dim GLHandle As Integer
  Dim GRHandle As Integer
  Dim y As Integer, z As Integer '8/10/04
  Dim Thisx As Integer
  ReDim TotDedAmt(1 To 50) As Double '8/19/04
  Dim CIDebit# '8/19/04
  Dim CICredit# '8/19/04
  ReDim FundArray(1 To 1) As String '8/19/04
  Dim ThisDesc$ '8/19/04
  Dim FirstFundSum As Boolean '8/19/04
  Dim FedTot As Double '8/19/04
  Dim MedTot As Double '8/19/04
  Dim SocTot As Double '8/19/04
  Dim StaTot As Double '8/19/04
  Dim RETTOT As Double '8/19/04
  Dim FSMTot As Double '8/19/04
  
  FirstFundSum = True
  FundCount = 0
  GLDebitTotal = 0
  GLCreditTotal = 0
  GLError = 0
  FEDWith$ = "FED Withholdings"      '    3
  MEDWith$ = "MED Withholdings"      '    4
  MEDMatLib$ = "MED Match Liab"      '    5
  SOCWith$ = "SOC Withholdings"      '    1
  SOCMatLib$ = "SOC Match Liab"      '    2
  RETWith$ = "RET Withholdings"      '    6
  RETMatLib$ = "RET Match Liab"      '    7
  STAWith$ = "STA Withholdings"      '    8

  MEDMat$ = "MED Match"
  SOCMat$ = "SOC Match"
  RETMat$ = "RET Match"              '


  ReDim SysRec(1) As RegDSysFileRecType
  ReDim PDR(1) As PeriodDefaultRecType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim EDesc(1) As String * 21
  ReDim Dash0(1) As String * 69
  LSet Dash0(1) = String$(69, "-")
  Dim FileHandle As Integer, Trans$
  Dim DedHandle As Integer, x As Integer
  Dim PDRHandle As Integer
  Dim SysHandle As Integer, Max As Integer
  Dim PPDefaultHandle As Integer
  Dim JFlag As Boolean, TotalAccts As Integer
  Dim GLIFRecName$, NumOfDeds As Integer

  FF$ = Chr$(12)
  TOTFEDTAX = 0
  TOTMEDTAX = 0
  TOTSOCTAX = 0
  TOTSTATAX = 0
  TOTRetTax = 0
  TOTMEDMat = 0
  TOTSOCMat = 0
  TOTRETMat = 0
  
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedHandle
  For x = 1 To 50
    Get DedHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedHandle
  
  OpenPPDefaultFile PDRHandle
  Get PDRHandle, 1, PDR(1)
  Close PDRHandle
  
  OpenSysFile SysHandle
  Get SysHandle, 1, SysRec(1)
  Close SysHandle
  
  Image3$ = "###,##0.00"

  FACnt = SysRec(1).AcctCnt

  NumOfWageAccts = UBound(DistbSumAccts)

  SplDedCnt = UBound(SplitDeds)

  'squeeze out all the "-" out of Acct numbers
  For cnt = 1 To NumOfWageAccts
    ReplaceString DistbSumAccts(cnt).Acct, "-", ""
  Next

  'change Period ending date to nicks format
  GLIFTDate$ = MakeRegDate(PDR(1).PEREND)

  ReplaceString GLIFTDate$, "-", "/"
  ReplaceString GLIFTDate$, "1994", "94"
  ReplaceString GLIFTDate$, "1995", "95"
  ReplaceString GLIFTDate$, "1996", "96"
  ReplaceString GLIFTDate$, "1997", "97"
  ReplaceString GLIFTDate$, "1998", "98"
  ReplaceString GLIFTDate$, "1999", "99"
  ReplaceString GLIFTDate$, "2000", "00"
  ReplaceString GLIFTDate$, "2001", "01"
  ReplaceString GLIFTDate$, "2002", "02"
  ReplaceString GLIFTDate$, "2003", "03"
  ReplaceString GLIFTDate$, "2004", "04"
  ReplaceString GLIFTDate$, "2005", "05"
  ReplaceString GLIFTDate$, "2006", "06"
  ReplaceString GLIFTDate$, "2007", "07"
  ReplaceString GLIFTDate$, "2008", "08"
  ReplaceString GLIFTDate$, "2009", "09"

  FedAcct$ = SysRec(1).Liab(1).Acct
  ReplaceString FedAcct$, "-", ""
  StaAcct$ = SysRec(1).Liab(2).Acct
  ReplaceString StaAcct$, "-", ""
  SocAcct$ = SysRec(1).Liab(3).Acct
  ReplaceString SocAcct$, "-", ""
  MedAcct$ = SysRec(1).Liab(4).Acct
  ReplaceString MedAcct$, "-", ""
  RetAcct$ = SysRec(1).Liab(5).Acct
  ReplaceString RetAcct$, "-", ""


  GLIFSource$ = GLIFTDate$
  ReplaceString GLIFSource$, "/", ""
  GLIFSource$ = "PR" + GLIFSource$

  GLIFSource$ = QPTrim$(GLIFSource$)
  GLIFTDate$ = QPTrim$(GLIFTDate$)
  
  GLIFArrayAdj = 56
  GLIFArrayAdj = SplDedCnt * GLIFArrayAdj

  ArraySize = NumOfWageAccts + GLIFArrayAdj + SplDedCnt

  ReDim GLIFRec(1 To ArraySize) As GLIFDataType14

  For cnt = 1 To NumOfWageAccts
    GLIFRec(cnt).TranAcct = DistbSumAccts(cnt).Acct
    GLIFRec(cnt).TranDate = GLIFTDate$
    GLIFRec(cnt).TranDesc = "Wages"
    GLIFRec(cnt).CrAmt = 0
    GLIFRec(cnt).DrAmt = DistbSumAccts(cnt).GrossPay
'    Debug.Print DistbSumAccts(cnt).GrossPay
    GLIFRec(cnt).Source = GLIFSource$
    GLIFRec(cnt).FromFlag = "W"
  Next
  NextAcct = cnt

'06-02-95 split deductions
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + FedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = FEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).FEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt
'look here
  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + StaAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = STAWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).STATAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + SocAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = SOCWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).SOCTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + MedAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = MEDWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).MEDTAX
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  NextAcct = NextAcct + cnt

  For cnt = 0 To (SplDedCnt - 1)
    GLIFRec(NextAcct + cnt).TranAcct = QPTrim$(SplitDeds(cnt + 1).Acct) + RetAcct$
    GLIFRec(NextAcct + cnt).TranDate = GLIFTDate$
    GLIFRec(NextAcct + cnt).TranDesc = RETWith$
    GLIFRec(NextAcct + cnt).CrAmt = SplitDeds(cnt + 1).RetTax
    GLIFRec(NextAcct + cnt).DrAmt = 0
    GLIFRec(NextAcct + cnt).Source = GLIFSource$
    GLIFRec(NextAcct + cnt).FromFlag = "X"
  Next

  'jumpthis:

  NextAcct = NextAcct + cnt

  'good to here ;maybe

  ReplaceString SysRec(1).CashAcct, "-", ""

  SysCash$ = QPTrim$(SysRec(1).CashAcct)

  If TotEIC# > 0 Then
    ReDim EICGLIFRec(1 To (SplDedCnt * 2)) As GLIFDataType14
    For cnt = 1 To SplDedCnt
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt).Acct) + FedAcct$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = 0
      EICGLIFRec(cnt).DrAmt = SplitDeds(cnt).EICPay
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next

    For cnt = SplDedCnt + 1 To SplDedCnt * 2
      EICGLIFRec(cnt).TranAcct = QPTrim$(SplitDeds(cnt - SplDedCnt).Acct) + SysCash$
      EICGLIFRec(cnt).TranDate = GLIFTDate$
      EICGLIFRec(cnt).TranDesc = "EIC Pmt"
      EICGLIFRec(cnt).CrAmt = SplitDeds(cnt - SplDedCnt).EICPay
      EICGLIFRec(cnt).DrAmt = 0
      EICGLIFRec(cnt).Source = GLIFSource$
      EICGLIFRec(cnt).FromFlag = "X"
    Next

  End If
  For ZZCnt = 0 To (SplDedCnt - 1)
    For cnt = 1 To NumOfDeds
      ReplaceString DedCodes(cnt).DCACCT1, "-", ""
      TAcct$ = QPTrim$(SplitDeds(ZZCnt + 1).Acct)
      TAcct$ = TAcct$ + QPTrim$(DedCodes(cnt).DCACCT1)
      GLIFRec(NextAcct + (cnt - 1)).TranAcct = TAcct$
      GLIFRec(NextAcct + (cnt - 1)).TranDate = GLIFTDate$
      GLIFRec(NextAcct + (cnt - 1)).TranDesc = QPTrim$(DedCodes(cnt).DCDESC1) '"Deductions"
      GLIFRec(NextAcct + (cnt - 1)).CrAmt = SplitDeds(ZZCnt + 1).DedAmts(cnt)
      GLIFRec(NextAcct + (cnt - 1)).DrAmt = 0
      GLIFRec(NextAcct + (cnt - 1)).Source = GLIFSource$
      GLIFRec(NextAcct + (cnt - 1)).FromFlag = "D"
    Next
    NextAcct = NextAcct + cnt
  Next

  ReDim DFunds$(1 To NumOfWageAccts)

  NumDFunds = 1
  DFunds$(1) = Left$(DistbSumAccts(1).Acct, 2)
  For cnt = 1 To NumOfWageAccts - 1
    If Left$(DistbSumAccts(cnt).Acct, 2) <> Left$(DistbSumAccts(cnt + 1).Acct, 2) Then
      NumDFunds = NumDFunds + 1 'counting the total number of funds.
      DFunds$(NumDFunds) = Left$(DistbSumAccts(cnt + 1).Acct, 2)
    End If
  Next

  ReDim Preserve DFunds$(1 To NumDFunds)

  ReDim DFund(1 To NumDFunds) As FundType

  For cnt = 1 To NumOfWageAccts
    CurrFund$ = Left$(DistbSumAccts(cnt).Acct, 2)
    For ThisFund = 1 To NumDFunds
      If CurrFund$ = DFunds$(ThisFund) Then
        DFund(ThisFund).FundCode = DFunds$(ThisFund)
        DFund(ThisFund).Debit = OldRound(DFund(ThisFund).Debit + DistbSumAccts(cnt).GrossPay)

        Exit For
      End If
    Next
  Next

  ''all gross pay by funds here!!

  ReDim CFunds$(1 To GLIFArrayAdj)              '17)

  FirstC = NumOfWageAccts + 1
  LastC = NumOfWageAccts + GLIFArrayAdj + 1     '17

  NumCFunds = 1
  CFunds$(1) = Left$(GLIFRec(FirstC).TranAcct, 2)

  For cnt = FirstC To LastC - 1
    If Left$(GLIFRec(cnt).TranAcct, 2) <> Left$(GLIFRec(cnt + 1).TranAcct, 2) Then
      If Len(QPTrim$(GLIFRec(cnt + 1).TranAcct)) Then
        NumCFunds = NumCFunds + 1               'counting the total number of funds.
        CFunds$(NumCFunds) = Left$(GLIFRec(cnt + 1).TranAcct, 2)
      End If
    End If
  Next
    
  ReDim Preserve CFunds$(1 To NumCFunds)
  ReDim CFund(1 To NumCFunds) As FundType
  For cnt = FirstC To LastC - 1
    CurrFund$ = Left$(GLIFRec(cnt).TranAcct, 2)
    For ThisFund = 1 To NumCFunds
      If CurrFund$ = CFunds$(ThisFund) Then
        CFund(ThisFund).FundCode = CFunds$(ThisFund)
        CFund(ThisFund).Credit = OldRound(CFund(ThisFund).Credit + GLIFRec(cnt).CrAmt)
        Exit For
      End If
    Next
  Next

  ''combine all funds in one array here
  TotalFunds = NumDFunds + NumCFunds            '+ 1
  ReDim AllFunds(1 To TotalFunds) As FundType
  ThisFund = 1
  For cnt = 1 To NumDFunds
    AllFunds(ThisFund) = DFund(cnt)
    ThisFund = ThisFund + 1
  Next

  For cnt = 1 To NumCFunds
    AllFunds(ThisFund) = CFund(cnt)
    ThisFund = ThisFund + 1
  Next

  '
  SortF AllFunds(), TotalFunds
  'combine Debits and Credits for same fund numbers

  First = 1
  Start = 1
  Last = TotalFunds
  Do
    Start = Start + 1
    For Cnt2 = Start To Last
      If AllFunds(First).FundCode = AllFunds(Cnt2).FundCode Then
        LSet AllFunds(Cnt2).FundCode = ""       'blank acct num as a flag
        AllFunds(First).Credit = OldRound#(AllFunds(First).Credit + AllFunds(Cnt2).Credit)

        AllFunds(Cnt2).Credit = 0
        AllFunds(First).Debit = OldRound#(AllFunds(First).Debit + AllFunds(Cnt2).Debit)

        AllFunds(Cnt2).Debit = 0
      End If
    Next
    First = First + 1
  Loop Until First >= Last      '

  'calc net difference for funds
  For cnt = 1 To TotalFunds
    If Len(QPTrim$(AllFunds(cnt).FundCode)) Then
      AllFunds(cnt).Net = OldRound#(AllFunds(cnt).Debit - AllFunds(cnt).Credit)
    End If
  Next

  'add cash acct number to fund numbers
  For cnt = 1 To TotalFunds
    FundCash$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(FundCash$) Then
      LSet AllFunds(cnt).FundCode = FundCash$ + SysCash$
    End If
  Next

  TotalGLIFS = NumOfWageAccts + GLIFArrayAdj + TotalFunds       '17

  ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14

  NextAcct = NumOfWageAccts + GLIFArrayAdj      '17

  For cnt = 1 To TotalFunds
    AcctNum$ = QPTrim$(AllFunds(cnt).FundCode)
    If Len(AcctNum$) Then
      NextAcct = NextAcct + 1
      GLIFRec(NextAcct).TranAcct = AcctNum$
      GLIFRec(NextAcct).TranDate = GLIFTDate$
      GLIFRec(NextAcct).TranDesc = "PR Net "
      GLIFRec(NextAcct).Source = GLIFSource$
      GLIFRec(NextAcct).FromFlag = "N"

      If AllFunds(cnt).Net > 0 Then
        GLIFRec(NextAcct).CrAmt = AllFunds(cnt).Net
        GLIFRec(NextAcct).DrAmt = 0
      ElseIf AllFunds(cnt).Net < 0 Then
        GLIFRec(NextAcct).DrAmt = Abs(AllFunds(cnt).Net)
        GLIFRec(NextAcct).CrAmt = 0
      End If
    End If
  Next

  'if using the imprest account then
  Select Case SysRec(1).USEIMP
  Case "I"      'was Y            'I C P
    TotalGLIFS = TotalGLIFS + 2
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    ReplaceString SysRec(1).ICRACCT, "-", ""
    GLIFRec(TotalGLIFS - 1).TranAcct = QPTrim$(SysRec(1).ICRACCT)
    GLIFRec(TotalGLIFS - 1).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS - 1).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS - 1).Source = GLIFSource$
    GLIFRec(TotalGLIFS - 1).FromFlag = "i"
    GLIFRec(TotalGLIFS - 1).DrAmt = Passed#(6)
    GLIFRec(TotalGLIFS - 1).CrAmt = 0
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "i"
    GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    GLIFRec(TotalGLIFS).DrAmt = 0
  Case "C"      'NEW Central Depository
    TotalGLIFS = TotalGLIFS + 1
    ReDim Preserve GLIFRec(1 To TotalGLIFS) As GLIFDataType14
    ReplaceString SysRec(1).IDRACCT, "-", ""
    GLIFRec(TotalGLIFS).TranAcct = QPTrim$(SysRec(1).IDRACCT)
    GLIFRec(TotalGLIFS).TranDate = GLIFTDate$
    GLIFRec(TotalGLIFS).TranDesc = "PPE " + GLIFTDate$
    GLIFRec(TotalGLIFS).Source = GLIFSource$
    GLIFRec(TotalGLIFS).FromFlag = "c"

    If TotEIC# > 0 Then
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6) - TotEIC#
    Else
      GLIFRec(TotalGLIFS).CrAmt = Passed#(6)
    End If

    GLIFRec(TotalGLIFS).DrAmt = 0

    ReplaceString SysRec(1).ICRACCT, "-", ""
    ReDim CDGLIFRec(1 To TotalFunds) As GLIFDataType14
    For cnt = 1 To TotalFunds
      If AllFunds(cnt).Net <> 0 Then
        CDGLIFRec(cnt).TranAcct = QPTrim$(SysRec(1).ICRACCT) + Left$(AllFunds(cnt).FundCode, 2)

        CDGLIFRec(cnt).TranDate = GLIFTDate$
        CDGLIFRec(cnt).TranDesc = "PPE " + GLIFTDate$
        CDGLIFRec(cnt).Source = GLIFSource$
        CDGLIFRec(cnt).FromFlag = "C"
        CDGLIFRec(cnt).DrAmt = AllFunds(cnt).Net                'Passed(6)
        CDGLIFRec(cnt).CrAmt = 0
      End If
    Next
  Case "P"

  End Select

  '*************END OF COMMON CODE SECTION

  If QPTrim$(SysRec(1).EXPMETHD) = "" Or SysRec(1).EXPMETHD = "0" Then
    GoTo WriteGLIFS
  End If

  ReDim SFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim MFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim RFGLIFRec(1 To TotalFunds) As GLIFDataType14
  ReDim SocGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim MedGLIFRec(1 To NumOfWageAccts) As GLIFDataType14
  ReDim RetGLIFRec(1 To NumOfWageAccts) As GLIFDataType14

  SOCEXP$ = QPTrim$(SysRec(1).SOCEXP)
  MEDEXP$ = QPTrim$(SysRec(1).MEDEXP)
  RETEXP$ = QPTrim$(SysRec(1).RETEXP)
  SOCLIAB$ = QPTrim$(SysRec(1).SOCLIAB)
  MEDLIAB$ = QPTrim$(SysRec(1).MEDLIAB)
  RETLIAB$ = QPTrim$(SysRec(1).RETLIAB)
  '
  ReplaceString SOCLIAB$, "-", ""
  ReplaceString MEDLIAB$, "-", ""
  ReplaceString RETLIAB$, "-", ""

  For cnt = 1 To NumOfWageAccts
    'social
    SocGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + SOCEXP$
    SocGLIFRec(cnt).TranDate = GLIFTDate$
    SocGLIFRec(cnt).TranDesc = SOCMat$
    SocGLIFRec(cnt).Source = GLIFSource$
    SocGLIFRec(cnt).FromFlag = "S"
    SocGLIFRec(cnt).CrAmt = 0
    SocGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATSocAmt
    'medicare
    MedGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + MEDEXP$
    MedGLIFRec(cnt).TranDate = GLIFTDate$
    MedGLIFRec(cnt).TranDesc = MEDMat$
    MedGLIFRec(cnt).Source = GLIFSource$
    MedGLIFRec(cnt).FromFlag = "M"
    MedGLIFRec(cnt).CrAmt = 0
    MedGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATMedAmt
    'retirment
    RetGLIFRec(cnt).TranAcct = Left$(DistbSumAccts(cnt).Acct, FACnt) + RETEXP$
    RetGLIFRec(cnt).TranDate = GLIFTDate$
    RetGLIFRec(cnt).TranDesc = RETMat$
    RetGLIFRec(cnt).Source = GLIFSource$
    RetGLIFRec(cnt).FromFlag = "R"
    RetGLIFRec(cnt).CrAmt = 0
    RetGLIFRec(cnt).DrAmt = DistbSumAccts(cnt).MATRetAmt
  Next

  For cnt = 1 To TotalFunds
    SFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + SOCLIAB$
    MFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + MEDLIAB$
    RFGLIFRec(cnt).TranAcct = Left$(AllFunds(cnt).FundCode, 2) + RETLIAB$
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(SFGLIFRec(Cnt2).TranAcct, 2) = Left$(SocGLIFRec(cnt).TranAcct, 2) Then
        SFGLIFRec(Cnt2).CrAmt = OldRound(SFGLIFRec(Cnt2).CrAmt + SocGLIFRec(cnt).DrAmt)
        SFGLIFRec(Cnt2).DrAmt = 0
        SFGLIFRec(Cnt2).TranDate = GLIFTDate$
        SFGLIFRec(Cnt2).TranDesc = SOCMatLib$
        SFGLIFRec(Cnt2).Source = GLIFSource$
        SFGLIFRec(Cnt2).FromFlag = "s"
      End If
    Next
  Next
  '
  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(MFGLIFRec(Cnt2).TranAcct, 2) = Left$(MedGLIFRec(cnt).TranAcct, 2) Then
        MFGLIFRec(Cnt2).CrAmt = OldRound(MFGLIFRec(Cnt2).CrAmt + MedGLIFRec(cnt).DrAmt)

        MFGLIFRec(Cnt2).DrAmt = 0
        MFGLIFRec(Cnt2).TranDate = GLIFTDate$
        MFGLIFRec(Cnt2).TranDesc = MEDMatLib$
        MFGLIFRec(Cnt2).Source = GLIFSource$
        MFGLIFRec(Cnt2).FromFlag = "m"
      End If
    Next
  Next

  For cnt = 1 To NumOfWageAccts
    For Cnt2 = 1 To TotalFunds
      If Left$(RFGLIFRec(Cnt2).TranAcct, 2) = Left$(RetGLIFRec(cnt).TranAcct, 2) Then
        RFGLIFRec(Cnt2).CrAmt = OldRound(RFGLIFRec(Cnt2).CrAmt + RetGLIFRec(cnt).DrAmt)
        RFGLIFRec(Cnt2).DrAmt = 0
        RFGLIFRec(Cnt2).TranDate = GLIFTDate$
        RFGLIFRec(Cnt2).TranDesc = RETMatLib$
        RFGLIFRec(Cnt2).Source = GLIFSource$
        RFGLIFRec(Cnt2).FromFlag = "r"
      End If
    Next
  Next
WriteGLIFS:

  GLIFRecLen = Len(GLIFRec(1))
  GLIFRecName$ = "TempIF.DAT"
  KillFile "TempIF.DAT"
  GHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GHandle Len = GLIFRecLen
  RecNo& = 1
  For cnt = 1 To TotalGLIFS
    If Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
'      If GLIFRec(cnt).DrAmt > 0 Or GLIFRec(cnt).CrAmt > 0 Then
      If GLIFRec(cnt).DrAmt <> 0 Or GLIFRec(cnt).CrAmt <> 0 Then '2/3/04 changed from line above
      'to allow for negative rounding adjustments in PCPrintPayRegisterST under FixSplit
        Put GHandle, RecNo&, GLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(SocGLIFRec(cnt).TranAcct)) Then
      If SocGLIFRec(cnt).DrAmt > 0 Or SocGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, SocGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(MedGLIFRec(cnt).TranAcct)) Then
      If MedGLIFRec(cnt).DrAmt > 0 Or MedGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, MedGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To NumOfWageAccts
    If Len(QPTrim$(RetGLIFRec(cnt).TranAcct)) Then
      If RetGLIFRec(cnt).DrAmt > 0 Or RetGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, RetGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(SFGLIFRec(cnt).TranAcct)) Then
      If SFGLIFRec(cnt).DrAmt > 0 Or SFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, SFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(MFGLIFRec(cnt).TranAcct)) Then
      If MFGLIFRec(cnt).DrAmt > 0 Or MFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, MFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  For cnt = 1 To TotalFunds
    If Len(QPTrim$(RFGLIFRec(cnt).TranAcct)) Then
      If RFGLIFRec(cnt).DrAmt > 0 Or RFGLIFRec(cnt).CrAmt > 0 Then
        Put GHandle, RecNo&, RFGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    End If
  Next

  If SysRec(1).USEIMP = "C" Then
    For cnt = 1 To TotalFunds
      If Len(QPTrim$(CDGLIFRec(cnt).TranAcct)) Then
        If CDGLIFRec(cnt).DrAmt <> 0 Or CDGLIFRec(cnt).CrAmt <> 0 Then
          Put GHandle, RecNo&, CDGLIFRec(cnt)
          RecNo& = RecNo& + 1
        End If
      End If
    Next
  End If

  'added EIC GLIF records if present 6/07/94
  If TotEIC# > 0 Then
    For cnt = 1 To (SplDedCnt * 2)
      If EICGLIFRec(cnt).DrAmt <> 0 Or EICGLIFRec(cnt).CrAmt <> 0 Then
        Put GHandle, RecNo&, EICGLIFRec(cnt)
        RecNo& = RecNo& + 1
      End If
    Next
  End If
  
  Close GHandle
  
  ReDim GLSAmt(1) As String * 11
  ReDim GLMAmt(1) As String * 11

'**** NEW
'  If Exist(QPTrim$(SysRec(1).CITIDIR) + "\" + JGLAcctIdxFile) Then
  If Exist(CurrCitiPath + "\" + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + "\" + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + JGLAcctIdxFile) Then
    GLIdxName$ = CurrCitiPath + JGLAcctIdxFile
    ReDim JGLIdxRec(1) As JGLAcctIdxType
    JFlag = True
  ElseIf Exist(CurrCitiPath + "\" + GLAcctIdxFile) Then
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + "\" + GLAcctIdxFile
  Else
    ReDim GLIdxRec(1) As GLAcctIdxType
'    GLIdxName$ = QPTrim$(SysRec(1).CITIDIR) + "\" + GLAcctIdxFile
    GLIdxName$ = CurrCitiPath + GLAcctIdxFile
  End If

  GLIdxRecLen = Len(JGLIdxRec(1))

  GLIFRecLen = Len(GLIFRec(1))

  TotalGLIFS = FileSize("TempIF.DAT") \ GLIFRecLen
  TotalAccts = FileSize(GLIdxName$) \ GLIdxRecLen

  If TotalGLIFS = 0 Then
    PrintGLRpt = False
    GoTo SkipGLRpt
  Else
    PrintGLRpt = True
  End If
  
  ReDim GLIFRec(1 To TotalGLIFS) As GLIFDataType14
  GRHandle = FreeFile
  Open GLIFRecName$ For Random Shared As GRHandle Len = GLIFRecLen
  
  For x = 1 To TotalGLIFS
    Get GRHandle, x, GLIFRec(x)
  Next x
'  SortG GLIFRec(), TotalGLIFS

  Do
    OutOfOrder = False                     'assume it's sorted
    For x = 1 To UBound(GLIFRec) - 1
      If GLIFRec(x).TranAcct > GLIFRec(x + 1).TranAcct Then
        SWAP GLIFRec(x), GLIFRec(x + 1)    'if we had to swap
        OutOfOrder = True                  'we're not done yet
      End If
    Next
  Loop While OutOfOrder

'04-05-96 Started Add summary to glif report
  For x = 1 To TotalGLIFS
    Put GRHandle, x, GLIFRec(x)
  Next x
  If TotalAccts = 0 Then
    Close
    GoTo SkipGLAccts
  End If
  Close GRHandle
  GLHandle = FreeFile
  Open GLIdxName$ For Random As GLHandle Len = GLIdxRecLen
'heredale
  ReDim JGoodAccts(1 To TotalAccts) As String * 16
  FrmShowPctComp.Label1 = "Printing Payroll Register"
  FrmShowPctComp.Show ' , Me
  For cnt = 1 To TotalAccts
    Get GLHandle, cnt, JGLIdxRec(1)
    ActualAccts = ActualAccts + 1
    ReplaceString JGLIdxRec(1).AcctNum, "-", ""
    JGoodAccts(ActualAccts) = JGLIdxRec(1).AcctNum
    FrmShowPctComp.ShowPctComp cnt, TotalAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close GLHandle
  TotalAccts = ActualAccts

SkipGLAccts:

  Max = 55
  Lines = 0
  RptName$ = "prrpts\PRGLIF.RPT"
  KillFile RptName$
  GHandle = FreeFile
  Open RptName$ For Output As GHandle
  RPTSetupPRN 14, GHandle
  THandle = FreeFile
  Open "TempIF.DAT" For Random As THandle Len = GLIFRecLen
  GoSub GLIFHeader

'**** NEW
  FrmShowPctComp.Label1 = "Printing G\L Interface Report"
  FrmShowPctComp.Show

  For cnt = 1 To TotalGLIFS
    FrmShowPctComp.ShowPctComp cnt, TotalGLIFS
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      Exit Sub
    End If
    If Lines >= Max Then
      Print #GHandle, FF$
      GoSub GLIFHeader
    End If

    TotDebit# = OldRound(TotDebit# + GLIFRec(cnt).DrAmt)
    TotCredit# = OldRound(TotCredit# + GLIFRec(cnt).CrAmt)

    '10-13-94 ** added check for valid GL Account numbers
    If TotalAccts > 0 Then
      AcctOk = False
      For Cnt2 = 1 To TotalAccts
        If InStr(JGoodAccts(Cnt2), QPTrim$(GLIFRec(cnt).TranAcct)) And Len(QPTrim$(JGoodAccts(Cnt2))) = Len(QPTrim$(GLIFRec(cnt).TranAcct)) Then
          AcctOk = True
          Exit For
        End If
      Next
      If Not AcctOk Then
        GLError = -1
        LSet GLIFRec(cnt).Fill = "ERR     "
        Put THandle%, cnt, GLIFRec(cnt)
      Else
        LSet GLIFRec(cnt).Fill = "        "
      End If
    Else 'added this code during Windows port 3/22/2002
      GLIFRec(cnt).Fill = ""
    End If

    'NoCheckAccts:
    RSet GLSAmt(1) = Using(Image3$, GLIFRec(cnt).DrAmt)
    RSet GLMAmt(1) = Using(Image3$, GLIFRec(cnt).CrAmt)
    LSet EDesc(1) = QPTrim$(GLIFRec(cnt).TranDesc)
    NoAcctNum = Len(QPTrim$(GLIFRec(cnt).TranAcct))
    '------added 8/19/04-------------------------
    
    ThisDesc$ = QPTrim$(GLIFRec(cnt).TranDesc)
    If FundCount = 0 Then
      FundCount = FundCount + 1
      ReDim FundArray(1 To FundCount) As String
      FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
      ReDim FedTaxByFund(1 To FundCount) As Double
      ReDim MedTaxByFund(1 To FundCount) As Double
      ReDim SocTaxByFund(1 To FundCount) As Double
      ReDim StaTaxByFund(1 To FundCount) As Double
      ReDim RetTaxByFund(1 To FundCount) As Double
      ReDim MedMatByFund(1 To FundCount) As Double
      ReDim SocMatByFund(1 To FundCount) As Double
      ReDim RetMatByFund(1 To FundCount) As Double
      ReDim DedAmtByFund(1 To 50, 1 To FundCount)
      Thisx = 1
   Else
     For x = 1 To FundCount
       If FundArray(x) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen) Then
         Thisx = x
         Exit For
       End If
     Next x
     If x > FundCount Then
       FundCount = FundCount + 1
       ReDim Preserve FundArray(1 To FundCount) As String
       FundArray(FundCount) = Mid(GLIFRec(cnt).TranAcct, 1, FundLen)
       ReDim Preserve FedTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedTaxByFund(1 To FundCount) As Double
       ReDim Preserve SocTaxByFund(1 To FundCount) As Double
       ReDim Preserve StaTaxByFund(1 To FundCount) As Double
       ReDim Preserve RetTaxByFund(1 To FundCount) As Double
       ReDim Preserve MedMatByFund(1 To FundCount) As Double
       ReDim Preserve SocMatByFund(1 To FundCount) As Double
       ReDim Preserve RetMatByFund(1 To FundCount) As Double
       ReDim Preserve DedAmtByFund(1 To 50, 1 To FundCount)
       Thisx = FundCount
     End If
   End If

    If ThisDesc = "FED Withholdings" Then
      FedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "SOC Withholdings" Then
      SocTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "MED Withholdings" Then
      MedTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "STA Withholdings" Then
      StaTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "RET Withholdings" Then
      RetTaxByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "SOC Match Liab" Then
      SocMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "MED Match Liab" Then
      MedMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    ElseIf ThisDesc = "RET Match Liab" Then
      RetMatByFund(Thisx) = GLIFRec(cnt).CrAmt
    Else
      For z = 1 To 50
        If ThisDesc = QPTrim$(DedCodes(z).DCDESC1) Then
          DedAmtByFund(z, Thisx) = GLIFRec(cnt).CrAmt
        End If
      Next z
    End If
    '----added 8/19/04--^^^^^^^^^^^
    If NoAcctNum > 0 Then
      If AcctOk = True Then
        LSet GLIFRec(cnt).Fill = " "
      End If
      Trans$ = QPTrim$(GLIFRec(cnt).TranAcct)
      LSet GLIFRec(cnt).TranAcct = " "
      LSet GLIFRec(cnt).TranAcct = Trans$
      Print #GHandle, "  " + GLIFRec(cnt).TranAcct + GLIFRec(cnt).Fill + EDesc(1) + GLSAmt(1) + "    " + GLMAmt(1)
      Lines = Lines + 1
    End If
  Next
  GoSub GLIFTotals
  Fund$ = Left$(QPTrim$(GLIFRec(1).TranAcct), 2)
  'added 8/19/04
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
    If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
      CIDebit = GLIFRec(1).DrAmt
      CICredit = GLIFRec(1).CrAmt
      GoTo NotNow1
    End If
  End If
  'added 8/19/04--^^^^
  FDebit# = GLIFRec(1).DrAmt
  FCredit# = GLIFRec(1).CrAmt
  
NotNow1:
  For cnt = 2 To TotalGLIFS
    NFund$ = Left$(QPTrim$(GLIFRec(cnt).TranAcct), 2)
    If NFund$ <> Fund$ Then
      GoSub PrintFundTotal
      Fund$ = NFund$
      If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
        If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
          CIDebit = GLIFRec(cnt).DrAmt
          CICredit = GLIFRec(cnt).CrAmt
          GoTo NotNow2
        End If
      End If
      FDebit# = GLIFRec(cnt).DrAmt
      FCredit# = GLIFRec(cnt).CrAmt
    Else
      If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
        If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
          CIDebit# = OldRound(CIDebit# + GLIFRec(cnt).DrAmt)
          CICredit# = OldRound(CICredit# + GLIFRec(cnt).CrAmt)
          GoTo NotNow2
        End If
      End If
      FDebit# = OldRound(FDebit# + GLIFRec(cnt).DrAmt)
      FCredit# = OldRound(FCredit# + GLIFRec(cnt).CrAmt)
    End If
NotNow2:
  Next

  GoSub PrintFundTotal
  GoSub PrintFundSummary
  Print #GHandle, FF$
  RPTSetupPRN 123, GHandle '7/24
  Close GHandle
  Close THandle
  
  If GLError <> -1 Then
    GLError = TotalAccts
  End If

SkipGLRpt:
  
  Exit Sub

PrintFundTotal:
    
  If FirstFundSum = True Then
    FirstFundSum = False
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Lines = 4
  End If
  For x = 1 To FundCount
    If FundArray(x) = Fund Then
      Thisx = x
      Exit For
    End If
  Next x
  
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then
    If Fund = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
      RSet GLSAmt(1) = Using(Image3$, CIDebit)
      RSet GLMAmt(1) = Using(Image3$, CICredit)
      GLDebitTotal = GLDebitTotal + GLSAmt(1)
      GLCreditTotal = GLCreditTotal + GLMAmt(1)
      GoTo CenImp
    End If
  End If
  RSet GLSAmt(1) = Using(Image3$, FDebit#)
  RSet GLMAmt(1) = Using(Image3$, FCredit#)
  LSet EDesc(1) = ""
  LSet GLIFRec(1).Fill = ""
  
CenImp:
'  If Lines >= Max Then
'    Print #GHandle, FF$
'    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
'    Print #GHandle, "G/L Interface Report Summary."
'    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
'    Print #GHandle, String$(80, "-")
'    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
'    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
'    Print #GHandle, String$(80, "-")
'    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
'    Print #GHandle,
'    Lines = 9
'  End If
  
  Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
  Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
  Print #GHandle, String$(80, "-")
  If QPTrim$(SysRec(1).USEIMP) = "C" Or QPTrim$(SysRec(1).USEIMP) = "I" Then 'might include imprest also
    If QPTrim$(Fund) = Mid(SysRec(1).ICRACCT, 1, FundLen) Then
      Lines = Lines + 3
      Return
    End If
  End If
  Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
  Print #GHandle,
  Lines = Lines + 5
  
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  
  TOTFEDTAX = OldRound(TOTFEDTAX + FedTaxByFund(Thisx))
  TOTMEDTAX = OldRound(TOTMEDTAX + MedTaxByFund(Thisx))
  TOTSOCTAX = OldRound(TOTSOCTAX + SocTaxByFund(Thisx))
  TOTSTATAX = OldRound(TOTSTATAX + StaTaxByFund(Thisx))
  TOTRetTax = OldRound(TOTRetTax + RetTaxByFund(Thisx))
  TOTMEDMat = OldRound(TOTMEDMat + MedMatByFund(Thisx))
  TOTSOCMat = OldRound(TOTSOCMat + SocMatByFund(Thisx))
  TOTRETMat = OldRound(TOTRETMat + RetMatByFund(Thisx))
  
  GLDebitTotal = GLDebitTotal + GLSAmt(1)
  GLCreditTotal = GLCreditTotal + GLMAmt(1)
  
  FedTot = FedTaxByFund(Thisx)
  MedTot = MedTaxByFund(Thisx) + MedMatByFund(Thisx)
  SocTot = SocTaxByFund(Thisx) + SocMatByFund(Thisx)
  StaTot = StaTaxByFund(Thisx)
  RETTOT = RetTaxByFund(Thisx) + RetMatByFund(Thisx)
  FSMTot = FedTot + MedTot + SocTot
  
  Print #GHandle, " Federal"; Tab(19); Using(Image3$, FedTaxByFund(Thisx)); Tab(71); Using$(Image3$, FedTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Social Security"; Tab(19); Using(Image3$, SocTaxByFund(Thisx)); Tab(46); Using(Image3$, SocMatByFund(Thisx)); Tab(71); Using$(Image3$, SocTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Medicare"; Tab(19); Using(Image3$, MedTaxByFund(Thisx)); Tab(46); Using(Image3$, MedMatByFund(Thisx)); Tab(71); Using$(Image3$, MedTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Sub Total"; Tab(71); Using(Image3$, FSMTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle,
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " State"; Tab(19); Using(Image3$, StaTaxByFund(Thisx)); Tab(71); Using$(Image3$, StaTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Retirement"; Tab(19); Using(Image3$, RetTaxByFund(Thisx)); Tab(46); Using(Image3$, RetMatByFund(Thisx)); Tab(71); Using$(Image3$, RETTOT)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, String$(80, "-")
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Lines = 7
  End If
  Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle,
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
    Print #GHandle,
    Lines = 9
  End If
  For z = 1 To 50
    TotDedAmt(z) = OldRound(TotDedAmt(z) + DedAmtByFund(z, Thisx))
    If QPTrim$(DedCodes(z).DCDESC1) <> "" Then
      Print #GHandle, Tab(2); QPTrim$(DedCodes(z).DCDESC1); Tab(19); Using(Image3$, DedAmtByFund(z, Thisx));
      Lines = Lines + 1
    End If
    If z = 50 Then Exit For
    TotDedAmt(z + 1) = OldRound(TotDedAmt(z + 1) + DedAmtByFund(z + 1, Thisx))
    If QPTrim$(DedCodes(z + 1).DCDESC1) <> "" Then
      Print #GHandle, Tab(51); QPTrim$(DedCodes(z + 1).DCDESC1); Tab(71); Using(Image3$, DedAmtByFund(z + 1, Thisx))
      z = z + 1
    End If
    If Lines >= Max Then
      Print #GHandle, FF$
      Print #GHandle, QPTrim$(Unit(1).UFEMPR)
      Print #GHandle, "G/L Interface Report Summary."
      Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
      Print #GHandle, String$(80, "-")
      Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
      Print #GHandle, "  Fund " + Fund; Tab(41); GLSAmt(1); Tab(70); GLMAmt(1)
      Print #GHandle, String$(80, "-")
      Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
      Print #GHandle,
      Lines = 9
    End If
  Next z
  Print #GHandle,
  Print #GHandle,
  Print #GHandle,
  Lines = Lines + 3
  If Lines >= Max - 4 Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle,
    Lines = 5
  End If
  Return
  
PrintFundSummary:
  If Lines >= Max - 10 Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Lines = 4
  End If
  
  MedTot = TOTMEDTAX + TOTMEDMat
  SocTot = TOTSOCTAX + TOTSOCMat
  RETTOT = TOTRetTax + TOTRETMat
  FSMTot = TOTFEDTAX + MedTot + SocTot
  
  Print #GHandle,
  Print #GHandle,
  Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
  Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
  Print #GHandle, String$(80, "-")
  Lines = Lines + 9
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Lines = 7
  End If
  Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
  Print #GHandle,
  Lines = Lines + 2
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Federal"; Tab(19); Using(Image3$, TOTFEDTAX); Tab(71); Using$(Image3$, TOTFEDTAX)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Social Security"; Tab(19); Using(Image3$, TOTSOCTAX); Tab(46); Using(Image3$, TOTSOCMat); Tab(71); Using$(Image3$, SocTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Medicare"; Tab(19); Using(Image3$, TOTMEDTAX); Tab(46); Using(Image3$, TOTMEDMat); Tab(71); Using$(Image3$, MedTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Sub Total"; Tab(71); Using(Image3$, FSMTot)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle,
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " State"; Tab(19); Using(Image3$, TOTSTATAX); Tab(71); Using$(Image3$, TOTSTATAX)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(17); "Withholdings"; Tab(40); "Matching Amounts"; Tab(75); "Totals"
    Print #GHandle,
    Lines = 9
  End If
  Print #GHandle, " Retirement"; Tab(19); Using(Image3$, TOTRetTax); Tab(46); Using(Image3$, TOTRETMat); Tab(71); Using$(Image3$, RETTOT)
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Lines = 6
  End If
  Print #GHandle, String$(80, "-")
  Lines = Lines + 1
  If Lines >= Max Then
    Print #GHandle, FF$
    Print #GHandle, QPTrim$(Unit(1).UFEMPR)
    Print #GHandle, "G/L Interface Report Summary."
    Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
    Print #GHandle, String$(80, "-")
    Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
    Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
    Print #GHandle, String$(80, "-")
    Lines = 7
  End If
  Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
  Print #GHandle,
  For z = 1 To 50
    If QPTrim$(DedCodes(z).DCDESC1) <> "" Then
      Print #GHandle, Tab(2); QPTrim$(DedCodes(z).DCDESC1); Tab(19); Using(Image3$, TotDedAmt(z));
      Lines = Lines + 1
    End If
    If z = 50 Then Exit For
    If QPTrim$(DedCodes(z + 1).DCDESC1) <> "" Then
      Print #GHandle, Tab(51); QPTrim$(DedCodes(z + 1).DCDESC1); Tab(71); Using(Image3$, TotDedAmt(z + 1))
      z = z + 1
    End If
    If Lines >= Max Then
      Print #GHandle, FF$
      Print #GHandle, QPTrim$(Unit(1).UFEMPR)
      Print #GHandle, "G/L Interface Report Summary."
      Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
      Print #GHandle, String$(80, "-")
      Print #GHandle, Tab(47); "DEBIT"; Tab(75); "CREDIT"
      Print #GHandle, "  Funds Grand Total"; Tab(42); Using(Image3$, GLDebitTotal); Tab(71); Using(Image3$, GLCreditTotal)
      Print #GHandle, String$(80, "-")
      Print #GHandle, Tab(2); "Deductions"; Tab(22); "Amounts"; Tab(51); "Deductions"; Tab(74); "Amounts"
      Print #GHandle,
      Lines = 8
    End If
  Next z
 
 Return
 
GLIFTotals:
  Print #GHandle,
  Print #GHandle, Dash0(1)
  Print #GHandle, "   Totals:                                  " + Using(Image3$, TotDebit#) + "     " + Using(Image3$, TotCredit#)
  If TotalAccts = 0 Then
    Print #GHandle,
    Print #GHandle, "  ERROR: G/L Accounts File NOT FOUND, or Invalid System Directory." '+ CrLf$
  End If
  Print #GHandle,
  Print #GHandle, ' "SubTotals:"
Return

GLIFHeader:
  Print #GHandle, QPTrim$(Unit(1).UFEMPR)
  Print #GHandle, "G/L Interface Report."
  Print #GHandle, "Period Ending: " + MakeRegDate(PDR(1).PEREND) ' + CrLf$
  Print #GHandle,
  Print #GHandle, "Account No.           Description                Debit         Credit"

  Print #GHandle, Dash0(1)
  Lines = 6
Return

End Sub

