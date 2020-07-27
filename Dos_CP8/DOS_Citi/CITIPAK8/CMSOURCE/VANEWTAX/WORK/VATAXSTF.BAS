DECLARE SUB VADisplayTaxScrn (ScrnName$)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB VAPrintPaymentJournal
  
  REDIM TaxSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TaxSetUp(1))
  FGetAH "TAXSETUP.DAT", TaxSetUp(1), TaxSetupLen, 1            'load it
  
  REDIM RERevAmts(1 TO 3) AS DOUBLE
  REDIM PPRevAmts(1 TO 7) AS DOUBLE
  
  REDIM RERevText$(1 TO 3)
  RERevText$(1) = "Tax Principle"
  RERevText$(2) = "Interest"
  RERevText$(3) = "Penalty"
  
  REDIM PPRevText$(1 TO 7)
  PPRevText$(1) = "Pers Prop Tax"
  PPRevText$(2) = "Mach/Tool Tax"
  PPRevText$(3) = "Merch Cap Tax"
  PPRevText$(4) = "Farm Equip Tax"
  PPRevText$(5) = "Mobile Hme Tax"
  PPRevText$(6) = "Interest"
  PPRevText$(7) = "Penalty"
  
  CursorOff
  Oper$ = QPTrim$(STR$(OperNum))
  FF$ = CHR$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = STRING$(80, "-")
  
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  
  PayJourName$ = "TAXPAY" + QPTrim$(STR$(OperNum)) + ".RPT"
  Header$ = "Utility Payment/Deposit Journal"
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  REDIM TaxPaymentRec(1) AS TaxPaymentRecType
  TaxPayRecLen = LEN(TaxPaymentRec(1))
  
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF CMOperRec(1).OperatorNumber = OperNum THEN
      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
      EXIT FOR
    END IF
  NEXT
  CLOSE CMFile
  
  IF Exist(TaxCPRFileName$) AND FileSize(TaxCPRFileName$) > 0 THEN
    PayOKFlag = True
  END IF
  
  IF NOT PayOKFlag THEN
    BlockClear
    VADisplayTaxScrn "NOPAYJUR"
    QPrintRC STR$(OperNum), 12, 34, 79
    WaitForAction
    GOTO ExitJournal
  END IF
  
  TotalRecs& = FileSize(TaxCPRFileName$) \ TaxPayRecLen
  BlockClear
  
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS RptHandle
  ShowProcessingScrn Header$
  
  GOSUB PrintRptHeader
  PayRecFile = FREEFILE
  OPEN TaxCPRFileName$ FOR RANDOM SHARED AS PayRecFile LEN = TaxPayRecLen
  NumOfRecs& = LOF(PayRecFile) \ TaxPayRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    GET #PayRecFile, Cnt&, TaxPaymentRec(1)
    DoneCnt = DoneCnt + 1
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader
    END IF
    IF TaxPaymentRec(1).CashAmt < 0 THEN TaxPaymentRec(1).CashAmt = 0
    IF TaxPaymentRec(1).ChkAmt < 0 THEN TaxPaymentRec(1).ChkAmt = 0
    
    PRINT #RptHandle, Num2Date(TaxPaymentRec(1).PayDate);
    PRINT #RptHandle, TAB(12); USING "#####"; TaxPaymentRec(1).CustAcct;
    PRINT #RptHandle, TaxPaymentRec(1).BillType;
    PRINT #RptHandle, TAB(19); LTRIM$(TaxPaymentRec(1).CustName);
    PRINT #RptHandle, TAB(44); USING "######.##"; TaxPaymentRec(1).CashAmt;
    PRINT #RptHandle, TAB(54); USING "######.##"; TaxPaymentRec(1).ChkAmt;
    PRINT #RptHandle, TAB(64); USING "######.##"; Round#(Round#(TaxPaymentRec(1).ChkAmt + TaxPaymentRec(1).CashAmt) - TaxPaymentRec(1).Change);
    PRINT #RptHandle, TAB(74); USING "####.##"; TaxPaymentRec(1).Change
    
    TotalCash# = Round#(TotalCash# + TaxPaymentRec(1).CashAmt)
    TotalCheck# = Round#(TotalCheck# + TaxPaymentRec(1).ChkAmt)
    TotalAmount# = Round#(TotalAmount# + TaxPaymentRec(1).AmtPaid)
    TotalChange# = Round#(TotalChange# + TaxPaymentRec(1).Change)
    TotalReceipts = TotalReceipts + 1
    LineCnt = LineCnt + 1
    IF TaxPaymentRec(1).BillType = "R" THEN
      FOR RCnt = 1 TO 3
        RERevAmts(RCnt) = Round#(RERevAmts(RCnt) + TaxPaymentRec(1).PaidOwed(RCnt).AmtPaid)
      NEXT
    ELSE
      FOR RCnt = 1 TO 7
        PPRevAmts(RCnt) = Round#(PPRevAmts(RCnt) + TaxPaymentRec(1).PaidOwed(RCnt).AmtPaid)
      NEXT
      
    END IF
    ShowPctComp DoneCnt, TotalRecs&
  NEXT
  
  GOSUB PrintRptEnding
  
  CLOSE
  
  LPTPORT = 1
  PrintRptFile Header$, PayJourName$, LPTPORT%, RetCode%, EntryPoint
  
  KillFile PayJourName$
  
ExitJournal:
  EXIT SUB
  
PrintRptHeader:
  Page = Page + 1
  PRINT #RptHandle, "Tax Payment Receipts Journal"
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(70); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "                                                                 Paid on"
  PRINT #RptHandle, " Date       Acct       Customer                 Cash     Check   Account  Change"
  PRINT #RptHandle, Dash1$
  LineCnt = 6
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "                  Totals: ";
  PRINT #RptHandle, TAB(44); USING "######.##"; TotalCash#;
  PRINT #RptHandle, TAB(54); USING "######.##"; TotalCheck#;
  PRINT #RptHandle, TAB(64); USING "######.##"; TotalAmount#;
  PRINT #RptHandle, TAB(74); USING "####.##"; TotalChange#
  
  PRINT #RptHandle, "Total Number of Receipts: "; USING "####,#"; TotalReceipts
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Tax Payment Receipts Journal"
  PRINT #RptHandle, "Source Summary"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Real Estate Receipts"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Source"; TAB(34); "Payments"
  PRINT #RptHandle, Dash1$
  REGTotal# = 0
  FOR RCnt = 1 TO 3
    PRINT #RptHandle, TAB(5); RERevText$(RCnt); TAB(30); USING "$$#####,#.##"; RERevAmts(RCnt)
    REGTotal# = Round#(REGTotal# + RERevAmts(RCnt))
    GTotal# = Round#(GTotal# + RERevAmts(RCnt))
  NEXT
  
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, " Real Estate Grand Total:"; TAB(39); USING "$$######,#.##"; REGTotal#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "Personal Property Receipts"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Source"; TAB(34); "Payments"
  PRINT #RptHandle, Dash1$
  REGTotal# = 0
  FOR RCnt = 1 TO 7
    PRINT #RptHandle, TAB(5); PPRevText$(RCnt); TAB(30); USING "$$#####,#.##"; PPRevAmts(RCnt)
    PPGTotal# = Round#(PPGTotal# + PPRevAmts(RCnt))
    GTotal# = Round#(GTotal# + PPRevAmts(RCnt))
  NEXT
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, " Personal Property Grand Total:"; TAB(39); USING "$$######,#.##"; PPGTotal#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Grand Total: All Receipts:"; TAB(39); USING "$$######,#.##"; GTotal#
  PRINT #RptHandle, FF$
RETURN
  
  
END SUB

