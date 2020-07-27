DEFINT A-Z
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (array())
DECLARE SUB SaveScrn (array())
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WaitForAction ()
DECLARE SUB HideCursor ()
DECLARE SUB ShowCursor ()
DECLARE SUB CursorOff ()

  CONST False = 0, True = NOT False

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'

   FF$ = CHR$(12)

SUB CustTaxReport
  
  SHARED Choice$()
  
  TheDate$ = DATE$
  Dash80$ = STRING$(80, "-")
  MaxLines = 60
  
  REDIM Choice$(0 TO 3, 0 TO 1)
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM TaxRates(1 TO 15) AS SINGLE
  REDIM RevTotals(1 TO 15) AS DOUBLE
  REDIM TaxAmt(1 TO 15) AS DOUBLE
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  TownName$ = UBSetup(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1
  
  FOR Cnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetup(1).Revenues(Cnt).RevName)
    IF LEN(TempRev$) = 0 THEN
      LastRev = Cnt - 1
      EXIT FOR
    ELSE
      RevText$(Cnt) = TempRev$
      TaxRates(Cnt) = UBSetup(1).Revenues(Cnt).TaxRate
      IF TaxRates(Cnt) > 0 THEN
        GotTaxFlag = True
      END IF
    END IF
  NEXT

  IF NOT GotTaxFlag THEN
    CursorOff
    HideCursor
    BlockClear
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "You do not have any taxes to report on.", 10, 21, -1
    QPrintRC "This report is ONLY for taxed revenues!", 11, 21, -1
    QPrintRC "Press any key to continue.", 13, 27, -1
    ShowCursor
    WaitForAction
    GOTO ExitTaxReport
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBTAXRPT"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  'Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  'Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = "Payments"
  Choice$(2, 0) = "Billings"
  
  ' Set Defaults as Anticpated Response
  FirstTime = True
  
  Action = 1
  BlockClear
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = TheDate$
      LSET Form$(2, 0) = TheDate$
      LSET Form$(4, 0) = "Y"
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckInfo
      IF ParmsOK THEN
        ExitFlag = True
      END IF
    CASE ESC
      ParmsOK = False
      ExitFlag = True
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF NOT ParmsOK THEN
    GOTO ExitTaxReport
  END IF
  
  BlockClear
  ShowProcessingScrn "Transaction Tax Report."
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBTAXRPT.RPT" FOR OUTPUT AS UBRpt
  
  GOSUB DoTaxRptHeader
  
  NumOfRecs& = LOF(UBCust) / UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    GET UBCust, Cnt&, UBCustRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, FF$
      GOSUB DoTaxRptHeader
    END IF
    Trans& = UBCustRec(1).LastTrans
    DO WHILE Trans& <> 0
      GET UBTrans, Trans&, UBTrans(1)
      IF UBTrans(1).TransType = TransType THEN
        IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) THEN
          IF UBTrans(1).TaxExempt <> "Y" THEN
            CustTax# = 0
            FOR RCnt = 1 TO LastRev
              IF TaxRates(RCnt) > 0 THEN
                Diff# = Round#(UBTrans(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                Tax# = Round#(UBTrans(1).RevAmt(RCnt) - Diff#)
                CustTax# = Round#(CustTax# + Tax#)
                TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
              END IF
            NEXT
            TTax# = Round#(TTax# + CustTax#)
            '*********************
            PRINT #UBRpt, Num2Date$(UBTrans(1).TransDate); TAB(11); USING "#####"; UBTrans(1).CustAcctNo;
            IF UBCustRec(1).DelFlag THEN
              PRINT #UBRpt, TAB(24); "*";
            END IF
            PRINT #UBRpt, TAB(25); LEFT$(UBCustRec(1).CustName, 33);
            PRINT #UBRpt, TAB(65); USING "$$#####,#.##"; CustTax#
            LineCnt = LineCnt + 1
            IF DetailFlag THEN
              FOR RCnt = 1 TO LastRev
                IF TaxRates(RCnt) > 0 THEN
                  Diff# = Round#(UBTrans(1).RevAmt(RCnt) / (1 + TaxRates(RCnt)))
                  Tax# = Round#(UBTrans(1).RevAmt(RCnt) - Diff#)
                  PRINT #UBRpt, RevText$(RCnt); TAB(16); USING "#####.##"; Tax#
                  LineCnt = LineCnt + 1
                END IF
              NEXT RCnt
              PRINT #UBRpt, Dash80$
              LineCnt = LineCnt + 1
            END IF
            IF LineCnt > MaxLines THEN
              PRINT #UBRpt, FF$
              GOSUB DoTaxRptHeader:
            END IF
          END IF
        END IF
      END IF
      
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
      Trans& = UBTrans(1).PrevTrans
    LOOP
    ShowPctComp Cnt&, NumOfRecs&
SkipThisOne:
  NEXT

  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Tax:"; TAB(65); USING "$$######,.##"; TTax#
  PRINT #UBRpt,
  PRINT #UBRpt, "Tax Breakdown:"
  FOR RCnt = 1 TO LastRev
    IF TaxRates(RCnt) > 0 THEN
      PRINT #UBRpt, TAB(5); RevText$(RCnt); TAB(20); USING "$$######,.##"; TaxAmt(RCnt)
    END IF
  NEXT
  PRINT #UBRpt,
  PRINT #UBRpt, "Report Parameters"
  PRINT #UBRpt, "     Report Type: "; Form$(3, 0)
  PRINT #UBRpt, "      Start Date: "; Form$(1, 0)
  PRINT #UBRpt, "     Ending Date: "; Form$(2, 0)
  PRINT #UBRpt, "          Detail: "; Form$(4, 0)

  CLOSE
  
  IF NOT AbortFlag THEN
    PrintRptFile "Transaction Tax Report.", "UBTAXRPT.RPT", 1, RetCode, EntryPoint
  END IF

ExitTaxReport:
  EXIT SUB

CheckInfo:
  BegDate = Date2Num(Form$(1, 0))
  EndDate = Date2Num(Form$(2, 0))
  Rpt$ = LEFT$(Form$(3, 0), 1)
  IF Form$(4, 0) = "Y" THEN
    DetailFlag = True
  ELSE
    DetailFlag = False
  END IF
  IF (BegDate < 0 OR EndDate < 0) OR (BegDate > EndDate) THEN
    Frm(1).FldNo = 1
    ErrCode = 1
    GOSUB ShowTaxParmErr
  ELSEIF LEN(QPTrim$(Form$(3, 0))) = 0 THEN
    Frm(1).FldNo = 3
    ErrCode = 2
    GOSUB ShowTaxParmErr
  ELSE
    ParmsOK = True
    SELECT CASE Rpt$
    CASE "P"
      TransType = TranBillPayment
    CASE "B"
      TransType = TranUtilityBill
    END SELECT
  END IF
  RETURN
  
ShowTaxParmErr:
  SaveScrn TempScrn()
  DisplayUBScrn "ERRSCRN1"
  SELECT CASE ErrCode
  CASE 1
    QPrintRC "Invalid Start/Ending Dates!", 10, 26, -1
  CASE 2
    QPrintRC "Invalid Report Type!", 10, 29, -1
  END SELECT
  QPrintRC "Correct and try again.", 13, 29, -1
  WaitForAction
  RestScrn TempScrn()
  Action = 1
  RETURN
  
DoTaxRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(30); "Transaction Tax Report"
  PRINT #UBRpt, TownName$; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; TheDate$
  PRINT #UBRpt, "  Date      Acct #             Customer Name                      Tax Amount"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
RETURN
  
  
END SUB

