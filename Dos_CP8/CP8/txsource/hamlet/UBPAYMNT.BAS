  DEFINT A-Z
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB DeletePayTrans (OPERNUM%)
DECLARE SUB DeleteDepTrans (OPERNUM%)
DECLARE SUB DeleteTransMenu (OPERNUM%)
DECLARE FUNCTION IsDeleted% (AcctNo&)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE SUB DelList (Items() AS ANY, Picked%(), NPicked%)
DECLARE SUB AddEditDeposit (OPERNUM%, PostDate$)
DECLARE SUB OperatorEntry (OperOKFlag%, OPERNUM%, PostDate$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB DisplayOperEntry ()
DECLARE SUB ValidateOperator (OPERNUM%, OperPassword$, OperOKFlag%)
DECLARE FUNCTION GetOperator% ()
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB SmallPause ()
DECLARE FUNCTION AskSavePayment% (ReceiptFlag$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB WaitForAction ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB AddEditPayment (OPERNUM, PostDate$)
DECLARE SUB PrintTransJournal (OPERNUM, PostDate$)
DECLARE SUB PostTransEntries (OPERNUM, PostDate$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
  'DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB ClearScrn ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Date2Num% (WhatDate$)
DECLARE FUNCTION Num2Date$ (WhatDate%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (FileHandle)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB ShowCustHistory (CUSTACCT&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB ButtonPress (WhatBtn, N, MooseButton, MCol, MRow)   ' ----- Check Mouse activity
DECLARE SUB GetCursor (a, b, NewStatus)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpRecLen%)
  
  '---- ASM subs
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.bi'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'newCUST.bi'
  '$INCLUDE: 'UBtrans.bi'
  '$INCLUDE: 'ubsetup.BI'
  '$INCLUDE: 'ubpaymnt.BI'
  '$INCLUDE: 'cmfiles.BI'
  
  ''$INCLUDE: 'init.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  CONST False = 0, True = NOT False
  
  'CALL TextCursor(14, 6)
  
  DIM SHARED OPERNUM AS INTEGER, OperOKFlag AS INTEGER
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  
  MChoice$(1) = "Payment Transaction Entry/Edit "
  MChoice$(2) = "Deposit Transaction Entry/Edit "
  MChoice$(3) = "Delete Payment/Deposit Transaction"
  MChoice$(4) = "Print Transaction Journal"
  MChoice$(5) = "Post Transaction Entries"
  MChoice$(6) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  
  BlockClear
  
  IF NOT WEnvTest% OR COMMAND$ = "TEST" THEN
    DO
      OperatorEntry OperOKFlag, OPERNUM, PostDate$
    LOOP WHILE OperOKFlag = 0
  ELSE
    OPERNUM = 1
    PostDate$ = DATE$
  END IF
  
  IF OperOKFlag < 0 THEN GOTO ExitPaymentProg
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Payments, Deposits Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditPayment OPERNUM, PostDate$
    CASE 2
      'AddEditDeposit OPERNUM, PostDate$
    CASE 3
      DeleteTransMenu OPERNUM
    CASE 4
      PrintTransJournal OPERNUM, PostDate$
    CASE 5
      PostTransEntries OPERNUM, PostDate$
    CASE 6
      HideCursor
      'ClearScrn
      END
    END SELECT
  LOOP
  
ExitPaymentProg:
  RUN "UBmenu"
  '  HideCursor
  '  ClearScrn
  '  END

SUB DeleteDepTrans (OPERNUM)
  
  PayFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  TPayFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".$$$"
  
  REDIM UBPaymentRec(1)   AS UBPaymentRecType
  
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  NumPayRecs = FileSize&(PayFileName$) \ UBPayRecLen
  
  TotalRecs& = NumPayRecs
  
  REDIM TransList(1 TO TotalRecs&)    AS FLen2
  REDIM Picked(1 TO TotalRecs&) AS INTEGER
  
  WhatCnt& = 1
  
  FOpenS PayFileName$, PHandle
  NumOfRecs& = FLof(PHandle) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      LSET TransList(WhatCnt&).V = UBPaymentRec(1).CUSTNAME
      MID$(TransList(WhatCnt&).V, 27) = "Deposit"
      MID$(TransList(WhatCnt&).V, 45) = FUsing$(STR$(UBPaymentRec(1).AMTPAID), "#####.##")
      'FUsing$ (Number$, Image$)
      WhatCnt& = WhatCnt& + 1
    NEXT
  END IF
  FClose PHandle
  PayListCnt& = NumOfRecs&
  
  MaxLen = 64
  Col = ((80 - MaxLen) \ 2) - 2
  
  BlockClear
  
  DelList TransList(), Picked(), NPicked
  
  IF NPicked > 0 THEN
    PFile = FREEFILE
    OPEN PayFileName$ FOR RANDOM SHARED AS PFile LEN = UBPayRecLen
    TFile = FREEFILE
    OPEN TPayFileName$ FOR RANDOM SHARED AS TFile LEN = UBPayRecLen
    FOR Cnt = 1 TO TotalRecs&
      GET PFile, Cnt, UBPaymentRec(1)
      SELECT CASE RIGHT$(TransList(Cnt).V, 1)
      CASE IS <> "û"
        PUT TFile, , UBPaymentRec(1)
      END SELECT
    NEXT
    CLOSE
    KillFile PayFileName$
    NAME TPayFileName$ AS PayFileName$
  END IF
  IF FileSize(PayFileName$) = 0 THEN
    KillFile PayFileName$
  END IF
END SUB

SUB PrintTransJournal (OPERNUM, PostDate$)
  
  REDIM RevText$(1 TO MaxRevsCnt)
  
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  
  REDIM UBSetUp(1) AS UBSetupRecType
  'LoadUBSetUpFile UBSetUp(), UBSetUpRecLen
  
  FOR Cnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUp(1).Revenues(Cnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      LastRev = Cnt - 1
      EXIT FOR
    ELSE
      RevText$(Cnt) = TempRev$
      TaxRates(Cnt) = UBSetUp(1).Revenues(Cnt).TaxRate
    END IF
  NEXT
  
  REDIM RevAmts(1 TO LastRev) AS DOUBLE
  REDIM DepRevs(1 TO LastRev) AS DOUBLE
  
  CursorOff
  Operator$ = QPTrim$(STR$(OPERNUM))
  FF$ = CHR$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = STRING$(107, "-")
  
  PayFileName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  DepFileName$ = "UBDEP" + QPTrim$(STR$(OPERNUM)) + ".DAT"
  PayJourName$ = "UBPAY" + QPTrim$(STR$(OPERNUM)) + ".RPT"
  
  Header$ = "Utility Payment/Deposit Journal"
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  
  REDIM UBPaymentRec(1) AS UBPaymentRecType
  UBPayRecLen = LEN(UBPaymentRec(1))
  
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  FOR Cnt = 1 TO NumRecs
    GET CMFile, Cnt, CMOperRec(1)
    IF CMOperRec(1).OperatorNumber = OPERNUM THEN
      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
      EXIT FOR
    END IF
  NEXT
  
  CLOSE CMFile
  
  IF Exist(PayFileName$) AND FileSize(PayFileName$) > 0 THEN
    PayOKFlag = True
  END IF
  IF Exist(DepFileName$) AND FileSize(DepFileName$) > 0 THEN
    DepOKFlag = True
  END IF
  
  IF NOT DepOKFlag AND NOT PayOKFlag THEN
    BlockClear
    'DisplayUBScrn "NOPAYJUR"
    QPrintRC STR$(OPERNUM), 12, 34, 79
    WaitForAction
    GOTO ExitJournal
  END IF
  
  TotalRecs& = (FileSize(DepFileName$) + FileSize(PayFileName$)) \ UBPayRecLen
  RptHandle = FREEFILE
  OPEN PayJourName$ FOR OUTPUT AS RptHandle
  ShowProcessingScrn Header$
  
  GOSUB PrintRptHeader
  IF PayOKFlag THEN
    FOpenS PayFileName$, PHandle
    NumOfRecs& = FLof(PHandle) \ UBPayRecLen
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      DoneCnt = DoneCnt + 1
      IF UBPaymentRec(1).TaxExempt = "Y" THEN
        TaxExempt = True
      ELSE
        TaxExempt = False
      END IF
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF UBPaymentRec(1).CASHAMT < 0 THEN UBPaymentRec(1).CASHAMT = 0
      IF UBPaymentRec(1).CHKAMT < 0 THEN UBPaymentRec(1).CHKAMT = 0
      
      PRINT #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
      PRINT #RptHandle, TAB(13); USING "#####"; UBPaymentRec(1).CUSTACCT;
      
      IF OPERNUM = 99 THEN
        Pmnt$ = " DFT"
      ELSE
        Pmnt$ = " PMT"
      END IF
      
      PRINT #RptHandle, TAB(20); UBPaymentRec(1).CUSTNAME; ; Pmnt$;
      
      PRINT #RptHandle, TAB(53); USING "#####,#.##"; UBPaymentRec(1).CASHAMT;
      PRINT #RptHandle, TAB(69); USING "#####,#.##"; UBPaymentRec(1).CHKAMT;
      
      PRINT #RptHandle, TAB(83); USING "#####,#.##"; Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).CHANGE);
      'PRINT #RptHandle, TAB(77); USING "#####,#.##"; UBPaymentRec(1).AMTOWED;
      PRINT #RptHandle, TAB(98); USING "#####,#.##"; UBPaymentRec(1).CHANGE
      
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      TotalAmount# = Round#(TotalAmount# + UBPaymentRec(1).AMTPAID)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).CHANGE)
      TotalReceipts = TotalReceipts + 1
      LineCnt = LineCnt + 1
      'findme
      'FOR RCnt = 1 TO LastRev
      '  RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
      'NEXT
      '******************
      FOR RCnt = 1 TO LastRev
        IF NOT TaxExempt THEN
          IF TaxRates(RCnt) > 0 THEN
            Diff# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 / (1 + TaxRates(RCnt)))
            Tax# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Diff#)
            TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + (UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Tax#))
          ELSE
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
          END IF
        ELSE
          RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        END IF
      NEXT
      '*********************
      
      ShowPctComp DoneCnt, TotalRecs&
    NEXT
    FClose PHandle
  END IF
  
  IF DepOKFlag THEN
    FOpenS DepFileName$, PHandle
    NumOfRecs& = FLof(PHandle) \ UBPayRecLen
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      DoneCnt = DoneCnt + 1
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF UBPaymentRec(1).CASHAMT < 0 THEN UBPaymentRec(1).CASHAMT = 0
      IF UBPaymentRec(1).CHKAMT < 0 THEN UBPaymentRec(1).CHKAMT = 0
      PRINT #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
      PRINT #RptHandle, TAB(13); USING "#####"; UBPaymentRec(1).CUSTACCT;
      PRINT #RptHandle, TAB(20); UBPaymentRec(1).CUSTNAME; ; " DEP";
      PRINT #RptHandle, TAB(53); USING "#####,#.##"; UBPaymentRec(1).CASHAMT;
      PRINT #RptHandle, TAB(69); USING "#####,#.##"; UBPaymentRec(1).CHKAMT;
      'PRINT #RptHandle, TAB(83); USING "#####,#.##"; UBPaymentRec(1).AMTPAID;
      PRINT #RptHandle, TAB(98); USING "#####,#.##"; UBPaymentRec(1).CHANGE
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).CHANGE)
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      '      TotalAmount# = Round#(TotalAmount# + UBPaymentRec(1).AMTPAID)
      TotalReceipts = TotalReceipts + 1
      LineCnt = LineCnt + 1
      FOR RCnt = 1 TO LastRev
        DepositTot# = Round#(DepositTot# + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        DepRevs(RCnt) = Round#(DepRevs(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
      NEXT
      ShowPctComp DoneCnt, TotalRecs&
    NEXT
    FClose PHandle
  END IF
  
  GOSUB PrintRptEnding
  
  CLOSE
  
  LPTPORT = 1
  
  PrintRptFile Header$, PayJourName$, LPTPORT%, RetCode%, EntryPoint
  
  'KillFile PayJourName$
  
ExitJournal:
  EXIT SUB
  
PrintRptHeader:
  Page = Page + 1
  PRINT #RptHandle, "Utility Payment/Deposit Receipts Journal"
  PRINT #RptHandle, "Posting Date: "; PostDate$
  PRINT #RptHandle, "    Operator: "; Operator$; TAB(89); "Page #"; Page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "       "; TAB(12); "        "; TAB(44); "           "; TAB(61); "            "; TAB(82); "Amount Paid"
  PRINT #RptHandle, " Date"; TAB(11); "Acct No      Customer"; TAB(52); "Cash Amount"; TAB(67); "Check Amount"; TAB(82); " on Account"; TAB(96); "Change Given"
  PRINT #RptHandle, Dash1$
  LineCnt = 6
  RETURN
  
PrintRptEnding:
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "                  Totals: ";
  PRINT #RptHandle, TAB(53); USING "#####,#.##"; TotalCash#;
  PRINT #RptHandle, TAB(69); USING "#####,#.##"; TotalCheck#;
  PRINT #RptHandle, TAB(83); USING "#####,#.##"; TotalAmount#;
  'PRINT #RptHandle, TAB(98); USING "#####,#.##"; Round#(Round#(TotalCash# + TotalCheck#) - TotalAmount#)
  PRINT #RptHandle, TAB(98); USING "#####,#.##"; TotalChange#
  
  PRINT #RptHandle, "Total Number of Receipts: "; USING "####,#"; TotalReceipts
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Utility Payment/Deposit Receipts Journal"
  PRINT #RptHandle, "Revenue Summary"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "    Revenue"; TAB(34); "Payments         Deposits               Tax"
  PRINT #RptHandle, Dash1$
  GTotal# = 0
  FOR RCnt = 1 TO LastRev
    PRINT #RptHandle, TAB(5); RevText$(RCnt); TAB(30); USING "$$#####,#.##"; RevAmts(RCnt); TAB(47); DepRevs(RCnt); TAB(65); TaxAmt(RCnt)
    GTotal# = Round#(GTotal# + RevAmts(RCnt))
    TTax# = Round#(TTax# + TaxAmt(RCnt))
  NEXT
  PRINT #RptHandle, Dash1$
  PRINT #RptHandle, "Revenue Totals:"; TAB(29); USING "$$######,#.##"; GTotal#; TAB(46); DepositTot#; TAB(64); TTax#
  PRINT #RptHandle,             '"Deposit Total:"; TAB(35); USING "$$#####,#.##"; DepositTot#
  PRINT #RptHandle, "   Grand Total:"; TAB(29); USING "$$######,#.##"; Round#(GTotal# + DepositTot# + TTax#)
  PRINT #RptHandle, FF$
  RETURN
  
END SUB

