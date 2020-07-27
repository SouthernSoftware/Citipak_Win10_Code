DEFINT A-Z
DECLARE SUB CursorOff ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)

DEFINT A-Z
SUB PrintPaymentJournal

  CLOSE

  REDIM PPD#(50), IPD#(50), CPD#(50), DPD#(50), Year(50)
  
  REDIM TaxCustRec(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCustRec(1))

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  REDIM TaxSetUp(1) AS TaxMasterType
  TaxSetupLen = LEN(TaxSetUp(1))

  REDIM PayListRec(1 TO 1)   AS PayListType
  PayListLen = LEN(PayListRec(1))

  REDIM TaxTranRec(1)        AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTranRec(1))

  FGetAH "TAXSETUP.DAT", TaxSetUp(1), TaxSetupLen, 1            'load it

 'Foxfire Flag Check
  FoxfireFlagChk = INSTR(TaxSetUp(1).Name, "FOXFIRE")
  
  REDIM RevAmts(1 TO 3) AS DOUBLE
  REDIM RevText$(1 TO 3)
  RevText$(1) = "Tax Principle"
  RevText$(2) = "Interest"
  RevText$(3) = "Collections"
  
  CursorOff
  Oper$ = QPTrim$(STR$(OperNum))
  FF$ = CHR$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = STRING$(80, "-")
  
  
  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers

  PayJourName$ = "TAXPAY" + QPTrim$(STR$(OperNum)) + ".RPT"
  Header$ = "Tax Payment Journal"
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  REDIM TaxPaymentRec(1) AS TaxPaymentRecType
  TaxPayRecLen = LEN(TaxPaymentRec(1))

  REDIM PayListRec(1 TO 1)   AS PayListType

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
    DisplayTaxScrn "NOPAYJUR"
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
  PayListFile = FREEFILE
  OPEN TaxLOPFileName$ FOR RANDOM AS PayListFile LEN = PayListLen
  TranFile = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TaxTranRecLen

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

    'Special for Foxfire  Print the Pin not the acct number for cust acct
    IF FoxfireFlagChk > 0 THEN
      CustFile = FREEFILE
      OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = TaxCustRecLen
      GET CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1)
      FirstPinRec& = TaxCustRec(1).FirstPropRec
      CLOSE CustFile
      IF FirstPinRec& > 0 THEN
        PropFile = FREEFILE
        OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
        GET PropFile, FirstPinRec&, PropRec(1)
        Acct$ = PropRec(1).RealPin
        CLOSE PropFile
      ELSE
        Acct$ = LTRIM$(STR$(TaxPaymentRec(1).CustAcct))
      END IF
     END IF

              
    PRINT #RptHandle, Num2Date(TaxPaymentRec(1).PayDate);
    IF FoxfireFlagChk > 0 THEN
      PRINT #RptHandle, TAB(12); LEFT$(Acct$, 6);
    ELSE
      PRINT #RptHandle, TAB(12); USING "#####"; TaxPaymentRec(1).CustAcct;
    END IF
    PRINT #RptHandle, TAB(19); TaxPaymentRec(1).CustName;
    PRINT #RptHandle, TAB(44); USING "######.##"; TaxPaymentRec(1).CashAmt;
    PRINT #RptHandle, TAB(54); USING "######.##"; TaxPaymentRec(1).ChkAmt;
    PRINT #RptHandle, TAB(64); USING "####.##"; TaxPaymentRec(1).DiscAmt;
    PRINT #RptHandle, TAB(72); USING "######.##"; Round#(Round#(TaxPaymentRec(1).ChkAmt + TaxPaymentRec(1).CashAmt + TaxPaymentRec(1).DiscAmt) - TaxPaymentRec(1).Change)
    TotalCash# = Round#(TotalCash# + TaxPaymentRec(1).CashAmt)
    TotalCheck# = Round#(TotalCheck# + TaxPaymentRec(1).ChkAmt)
    TotalAmount# = Round#(TotalAmount# + TaxPaymentRec(1).AmtPaid)
    TotalDisc# = Round#(TotalDisc# + TaxPaymentRec(1).DiscAmt)
    TotalChange# = Round#(TotalChange# + TaxPaymentRec(1).Change)
    TotalReceipts = TotalReceipts + 1
    LineCnt = LineCnt + 1
    FOR RCnt = 1 TO 3
      RevAmts(RCnt) = Round#(RevAmts(RCnt) + TaxPaymentRec(1).PaidOwed(RCnt).AmtPaid)
    NEXT

   'Special for Subtotaling by year

   ThisListRec& = TaxPaymentRec(1).LastPayRec
    DO WHILE ThisListRec& > 0
      GET #PayListFile, ThisListRec&, PayListRec(1)
      'get paylist rec

      GET #TranFile, PayListRec(1).BillRec, TaxTranRec(1)
      'get bill trans this payrec is for
      TotalPaid# = Round#(PayListRec(1).Principle1 + PayListRec(1).Interest1 + PayListRec(1).Collection)
      IF TotalPaid# = 0 THEN
        GOTO SkipThisRec1
      END IF

      PPD# = Round#(PayListRec(1).Principle1)
      IPD# = Round#(PayListRec(1).Interest1)
      CPD# = Round#(PayListRec(1).Collection)
      DPD# = Round#(TaxPaymentRec(1).DiscAmt)
      Year = TaxTranRec(1).TaxYear

      IF YCnt& = 0 THEN
       YCnt& = 1
       PPD#(YCnt&) = PPD#
       IPD#(YCnt&) = IPD#
       CPD#(YCnt&) = CPD#
       DPD#(YCnt&) = DPD#
       Year(YCnt&) = Year
       GOTO SkipThisRec1
      END IF

      FOR YS& = 1 TO YCnt&
      IF Year = Year(YS&) THEN
       PPD#(YS&) = PPD#(YS&) + PPD#
       IPD#(YS&) = IPD#(YS&) + IPD#
       CPD#(YS&) = CPD#(YS&) + CPD#
       DPD#(YS&) = DPD#(YS&) + DPD#
       GOTO SkipThisRec1
      END IF
      NEXT YS&
      YCnt& = YCnt& + 1
      PPD#(YCnt&) = PPD#
      IPD#(YCnt&) = IPD#
      CPD#(YCnt&) = CPD#
      DPD#(YCnt&) = DPD#
      Year(YCnt&) = Year
SkipThisRec1:
      ThisListRec& = PayListRec(1).PrevListRec
     LOOP
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
  PRINT #RptHandle, "                                                                         Total"
  PRINT #RptHandle, "                                                                         Amount"
  IF FoxfireFlagChk > 0 THEN
   PRINT #RptHandle, " Date       PIN#       Customer                 Cash     Check  Disc't  Credited"
  ELSE
   PRINT #RptHandle, " Date       Acct       Customer                 Cash     Check  Disc't  Credited"
  END IF
  PRINT #RptHandle, Dash1$
  LineCnt = 6
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "                  Totals: ";
  PRINT #RptHandle, TAB(44); USING "######.##"; TotalCash#;
  PRINT #RptHandle, TAB(54); USING "######.##"; TotalCheck#;
  PRINT #RptHandle, TAB(64); USING "####.##"; TotalDisc#;
  PRINT #RptHandle, TAB(72); USING "######.##"; TotalAmount#
  PRINT #RptHandle, "Total Change Returned: "; USING "######,#.##"; TotalChange#
  
  PRINT #RptHandle, "Total Number of Receipts: "; USING "####,#"; TotalReceipts
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Tax Payment Receipts Journal"
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(70); "Page #"; Page
  PRINT #RptHandle, "Source Summary"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "     Source"; TAB(34); "Payments/Discounts"
  PRINT #RptHandle, Dash1$
  GTotal# = 0
  FOR RCnt = 1 TO 3
    PRINT #RptHandle, TAB(5); RevText$(RCnt); TAB(30); USING "$$#####,#.##"; RevAmts(RCnt)
    GTotal# = Round#(GTotal# + RevAmts(RCnt))
  NEXT
  
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, " Grand Total:"; TAB(29); USING "$$######,#.##"; GTotal#
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Tax Payment Receipts Journal : Yearly Breakdown"
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(70); "Page #"; Page
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(59); "Discounts"; TAB(70); "Payments"
  PRINT #RptHandle, "Tax Year"; TAB(15); "Principle Pd"; TAB(30); "Interest Pd"; TAB(45); "Collect Pd"; TAB(60); "Allowed"; TAB(70); "Less Disc"
  PRINT #RptHandle, Dash1$
  FOR TestYear = 2020 TO 1980 STEP -1
  FOR TYear& = 1 TO YCnt&
  IF Year(TYear&) = TestYear THEN
  PRINT #RptHandle, Year(TYear&);
   PRINT #RptHandle, TAB(15); USING "$$#####,#.##"; PPD#(TYear&);
   PRINT #RptHandle, TAB(28); USING "$$#####,#.##"; IPD#(TYear&);
   PRINT #RptHandle, TAB(41); USING "$$#####,#.##"; CPD#(TYear&);
   PRINT #RptHandle, TAB(54); USING "$$#####,#.##"; DPD#(TYear&);
   PRINT #RptHandle, TAB(67); USING "$$#####,#.##"; PPD#(TYear&) + IPD#(TYear&) + CPD#(TYear&) - DPD#(TYear&)
   TPPD# = TPPD# + PPD#(TYear&)
   TIPD# = TIPD# + IPD#(TYear&)
   TCPD# = TCPD# + CPD#(TYear&)
   TDPD# = TDPD# + DPD#(TYear&)
   END IF
  NEXT TYear&
  NEXT TestYear
   PRINT #RptHandle, Dash1$
   PRINT #RptHandle, "Totals";
   PRINT #RptHandle, TAB(15); USING "$$#####,#.##"; TPPD#;
   PRINT #RptHandle, TAB(28); USING "$$#####,#.##"; TIPD#;
   PRINT #RptHandle, TAB(41); USING "$$#####,#.##"; TCPD#;
   PRINT #RptHandle, TAB(54); USING "$$#####,#.##"; TDPD#;
   PRINT #RptHandle, TAB(67); USING "$$#####,#.##"; TPPD# + TIPD# + TCPD# - TDPD#
   PRINT #RptHandle, FF$
RETURN
  
END SUB

