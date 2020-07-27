DECLARE SUB ShowProcessing (Cnt&, NumOfRecs&)
DECLARE SUB ShowWarning ()

DEFINT A-Z
SUB FixRunBalance

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM TranRecs(0 TO 1) AS ArcTranType

  ShowWarning
  PRINT "   RECALCULATE RUNNING BALANCES."
  Ok$ = GetProceed$

  SELECT CASE Ok$
  CASE "Y"
    PRINT Ok$
    GOSUB GOCalcRunBal
'    IF ThruYear <> -1 THEN
'      PRINT
'      PRINT
'      PRINT "   Transaction History Archived."
'    END IF
  CASE ELSE
    PRINT "N"
    PRINT
    PRINT "   RECALCULATE BALANCES ABORTED!"
  END SELECT
EXIT SUB
'
GOCalcRunBal:
  PRINT
  
  UBTran = FREEFILE
  OPEN "UBTrans.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  Cnt& = 1000
  'FOR Cnt& = 1 TO NumOfRecs&
    ShowProcessing Cnt&, NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    TranCnt = 0
    ThisTran& = UBCustRec(1).LastTrans
    REDIM TranRecs(0 TO 1) AS ArcTranType
    DO WHILE ThisTran& > 0
      TranCnt = TranCnt + 1
      REDIM PRESERVE TranRecs(0 TO TranCnt)  AS ArcTranType
      GET UBTran, ThisTran&, UBTranRec(1)
      TranRecs(TranCnt).RecNo = ThisTran&
      TranRecs(TranCnt).RunBalance = UBTranRec(1).RunBalance
      IF UBTranRec(1).TransType > 100 THEN
        TranRecs(TranCnt).TRType = UBTranRec(1).TransType - 100
      ELSE
        TranRecs(TranCnt).TRType = UBTranRec(1).TransType
      END IF
      ThisTran& = UBTranRec(1).PrevTrans
    LOOP
    IF TranCnt > 0 THEN
      ActualBal# = 0
      FOR Cnt = TranCnt TO 1 STEP -1
        GET UBTran, TranRecs(Cnt).RecNo, UBTranRec(1)
        IF Cnt = TranCnt THEN
          TRunBal# = TranRecs(Cnt).RunBalance
        END IF
        SELECT CASE TranRecs(Cnt).TRType
        CASE TranUtilityBill           '   1=Utility bill
          ActualBal# = uRound#(ActualBal# + UBTranRec(1).TransAmt)
          IF TranRecs(Cnt).RunBalance <> RunBal# THEN

        CASE TranLateCharge            '   2=late charge      'NOT USED

        CASE TranReconnectFee          '   3=reconnect fee    'NOT USED

        CASE TranBillPayment           '   4=Bill Payment

        CASE TranAppliedDeposit        '   5=Applied Deposit

        CASE TranPenaltyCharge         '   6=Penalty Charge

        CASE TranDepositPayment        '   7=Deposit Payment

        CASE TranDraftPayment          '   8=Draft Payment

        CASE TranRefundDeposit         '   9=Refund Deposit

        CASE TranBeginBalance          '  10=Beginning Balance

        CASE TranUpwardAdjustment      '  11=Bill Adjustments

        CASE TranDownwardAdjustment    '  12=Bill Adjustments

        END SELECT
      '  LPRINT TranRecs(Cnt).TRType, TranRecs(Cnt).RunBalance
      

      NEXT
    END IF
    CLOSE
    END



'    'first make sure they have transactions greater than this date range
'
'    CanDo = False
'    FOR TCnt = 1 TO TranCnt
'      IF TranRecs(TCnt).TRDate > EndDate THEN
'        CanDo = True
'        EXIT FOR
'      END IF
'    NEXT
'    IF CanDo THEN
'      FOR TCnt = 1 TO TranCnt
'        IF TranRecs(TCnt).TRDate <= EndDate THEN
'          SELECT CASE TranRecs(TCnt).TRType
'          'keep these transaction anyway
'          CASE TranAppliedDeposit, TranDepositPayment, TranRefundDeposit
'            GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'            PUT NUBTran, , UBTranRec(1)
'          CASE ELSE
'            Arked& = Arked& + 1
'          END SELECT
'        ELSE
'          GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'          PUT NUBTran, , UBTranRec(1)
'        END IF
'      NEXT
'    ELSE 'no trans greater then date; copy all of them
'      FOR TCnt = 1 TO TranCnt
'        GET UBTran, TranRecs(TCnt).RecNo, UBTranRec(1)
'        PUT NUBTran, , UBTranRec(1)
'      NEXT
'    END IF
'    'PUT UBFile, Cnt&, UBCustRec(1)
'  NEXT
'  CLOSE
'  PRINT
'  PRINT
'  PRINT "   Removed:"; Arked&; " Tranactions"
CalcRunBalExit:
'
'RETURN

END SUB

