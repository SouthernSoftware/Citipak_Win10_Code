DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB oPostFinalBillTrans
  
  UBLog "IN: POST FINAL"
  
  REDIM DepRev(1 TO 15) AS DOUBLE
  REDIM UBTempDepTran(1) AS UBTransRecType
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "CLEVELAND") THEN
    CleveFlag = True
    UBLog "POST FINAL:  CLEVELAND Detected "
  END IF
  
  IF NOT Exist(UBFinBillsFile) THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NON2POST"
    WaitForAction
    GOTO ExitBillPost:
  END IF
  
  IF NOT Exist("UBFBILLS.PRN") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitBillPost:
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "PSTBILLS"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OKFlag OR ExitFlag
  
  IF ExitFlag THEN
    BlockClear
    CursorOff
    DisplayUBScrn "POSTCAN"
    WaitForAction
    GOTO ExitBillPost
  END IF
  
  REDIM UBBillRec(1) AS UBTransRecType
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  
  UBBillRecLen = LEN(UBBillRec(1))
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBBillRecLen
  
  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen
  ShowProcessingScrn "Posting Final Billing Transactions"
  '*****************************************************
  
  FOR BillCnt = 1 TO NumOfBillRecs
    GET UBBill, BillCnt, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN             'AND UBBillRec(1).TransAmt > 0 THEN
      GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      UBCustRec(1).Status = "B"
      UBCustRec(1).PrevBalance = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      
      IF UBBillRec(1).ApplyDepFlag = "Y" THEN
        IF CleveFlag THEN
          GOSUB ProcCleveDeposit
        ELSE
          GOSUB ProcCustDeposit
        END IF
      ELSE
        FOR RevCnt = 1 TO MaxRevsCnt
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
        NEXT
      END IF
      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
      FOR MtrCnt = 1 TO 7
        CubMtr = False
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          IF UBCustRec(1).LocMeters(MtrCnt).MTRUnit = "C" THEN
            CubMtr = True
          END IF
          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
          IF ReadAmt& < 0 THEN  'Meter rolled over or, misread
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
          END IF
          IF CubMtr THEN
            ReadAmt& = ReadAmt& * 7.481
          END IF
          UBCustRec(1).LocMeters(MtrCnt).AvgUse = Round(UBCustRec(1).LocMeters(MtrCnt).AvgUse + ReadAmt&)
          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
          '050697 Fixed current reading not being move to previous
          UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBCustRec(1).LocMeters(MtrCnt).CurRead
        END IF
      NEXT
      
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBBillRec(1).PrevTrans = PrevLastTrans&
      NextTransRec& = (LOF(UBTran) \ UBBillRecLen) + 1          'point at next trans
      PUT UBTran, NextTransRec&, UBBillRec(1)
      UBCustRec(1).LastTrans = NextTransRec&
      
      'detach the new vacant rec from this customer
      IF UBCustRec(1).OldRec > 0 THEN
        UBLog "POST FINAL: DETACHED OLD ACCT:" + STR$(UBCustRec(1).OldRec)
        UBCustRec(1).OldRec = 0
      END IF
      PUT UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)
      '040997 added Transaction to show customers applied deposit
      IF DepAppliedFlag THEN
        GOSUB MakeAppDepTrans
      END IF
    END IF
    ShowPctComp BillCnt, NumOfBillRecs
  NEXT
  
  CLOSE
  
  KillFile UBFinBillsFile
  KillFile "UBFBILLS.PRN"
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
  
ExitBillPost:
  UBLog "OUT: POST FINAL"
  EXIT SUB
  
MakeAppDepTrans:
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  UBTransRec(1).TransDate = UBBillRec(1).TransDate
  'UBTransRec(1)CustLocation = UBBillRec(1).CustAcctNo
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = UBBillRec(1).CustAcctNo
  UBTransRec(1).TransAmt = DepTranAmt#
  '091198 Changed to put original deposit amounts in revenue source
  FOR Cnt = 1 TO 15
    UBTransRec(1).RevAmt(Cnt) = DepRev(Cnt)
  NEXT
  
  UBTransRec(1).TransDesc = "Applied Deposit"
  UBTransRec(1).TransType = TranAppliedDeposit
  UBTransRec(1).RunBalance = Round#((UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance) - ABS(DepTranAmt#))
  UBCustRec(1).DepositAmt = 0
  UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance - ABS(DepTranAmt#))
  
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  
  IF Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
    IF UBCustRec(1).Status = "B" THEN
      CustChCnt = CustChCnt + 1
      UBLog "POST FINAL: SET CUST STATUS=I ACCT:" + STR$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    END IF
  END IF
  NextTransRec& = (LOF(UBTran) \ UBTransRecLen) + 1             'point at next trans
  PUT UBTran, NextTransRec&, UBTransRec(1)
  UBCustRec(1).LastTrans = NextTransRec&
  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
  UBLog "POST FINAL: DEP APPLIED TRANS:" + STR$(NextTransRec&)
RETURN
  
  
ProcCleveDeposit:
  FOR LLCnt = 1 TO 15: DepRev(LLCnt) = 0: NEXT LLCnt
    DepAppliedFlag = False
    DepTranAmt# = -UBCustRec(1).DepositAmt
    DepositAmt# = UBCustRec(1).DepositAmt
    IF DepositAmt# = 0 THEN
      GOTO NoDepReturn
    END IF
    ThisTran& = UBCustRec(1).LastTrans
    DO WHILE ThisTran& > 0
      GET UBTran, ThisTran&, UBTempDepTran(1)
      IF UBTempDepTran(1).TransType = TranDepositPayment THEN
        FOR DZCnt = 1 TO 15
          DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
        NEXT
      END IF
      ThisTran& = UBTempDepTran(1).PrevTrans
    LOOP
    
    FOR RevCnt = 1 TO MaxRevsCnt - 1
      UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
      IF DepRev(RevCnt) > 0 THEN
        DepAppliedFlag = True
        IF UBCustRec(1).CurrRevAmts(RevCnt) < DepRev(RevCnt) THEN
          DepRev(RevCnt) = Round#(DepRev(RevCnt) - UBCustRec(1).CurrRevAmts(RevCnt))
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
        ELSEIF UBCustRec(1).CurrRevAmts(RevCnt) > DepRev(RevCnt) THEN
          UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - DepRev(RevCnt))
          DepRev(RevCnt) = 0
        ELSE    'the deposit and the revenue are equal
          UBCustRec(1).CurrRevAmts(RevCnt) = 0
          DepRev(RevCnt) = 0
        END IF
      END IF
    NEXT
    
    'If there was any deposit left after applying to the cust rev totals
    FOR RevCnt = 1 TO MaxRevsCnt - 1
      IF DepRev(RevCnt) > 0 THEN
        UBCustRec(1).CurrRevAmts(RevCnt) = -DepRev(RevCnt)
      END IF
    NEXT
    UBCustRec(1).DepositAmt = 0
    
NoDepReturn:
    
    RETURN
    
ProcCustDeposit:
    FOR LLCnt = 1 TO 15: DepRev(LLCnt) = 0: NEXT LLCnt
      DepAppliedFlag = False
      DepositAmt# = UBCustRec(1).DepositAmt
      DepTranAmt# = -UBCustRec(1).DepositAmt
      FOR RevCnt = 1 TO MaxRevsCnt - 1
        UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
        IF DepositAmt# > 0 THEN
          DepAppliedFlag = True
          IF UBCustRec(1).CurrRevAmts(RevCnt) < DepositAmt# THEN
            DepositAmt# = Round(DepositAmt# - UBCustRec(1).CurrRevAmts(RevCnt))
            UBCustRec(1).CurrRevAmts(RevCnt) = 0
          ELSEIF UBCustRec(1).CurrRevAmts(RevCnt) > DepositAmt# THEN
            UBCustRec(1).CurrRevAmts(RevCnt) = Round(UBCustRec(1).CurrRevAmts(RevCnt) - DepositAmt#)
            DepositAmt# = 0
          ELSE  'the deposit and the revenue are equal
            UBCustRec(1).CurrRevAmts(RevCnt) = 0
            DepositAmt# = 0
          END IF
        END IF
      NEXT
      'If there was any deposit left after applying to the cust rev totals
      IF DepositAmt# > 0 THEN
        UBCustRec(1).CurrRevAmts(1) = -DepositAmt#
      END IF
      UBCustRec(1).DepositAmt = 0
      
      ThisTran& = UBCustRec(1).LastTrans
      DO WHILE ThisTran& > 0
        GET UBTran, ThisTran&, UBTempDepTran(1)
        IF UBTempDepTran(1).TransType = TranDepositPayment THEN
          FOR DZCnt = 1 TO 15
            DepRev(DZCnt) = Round#(DepRev(DZCnt) + UBTempDepTran(1).RevAmt(DZCnt))
          NEXT
        END IF
        ThisTran& = UBTempDepTran(1).PrevTrans
      LOOP
RETURN
END SUB

