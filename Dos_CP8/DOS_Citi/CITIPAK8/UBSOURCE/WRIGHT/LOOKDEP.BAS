DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'Newcust.BI'

  TranDate = Date2Num(DATE$)
  'PastDate = Date2Num("11-15-1996")

  CLS
  LOCATE 3, 1, 0
  PRINT "Checking Customer Balances. . ."
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBTranRecLen = LEN(UBTranRec(1))
  UBCustRecLen = LEN(UBCustRec(1))

  'NAME "UBTRANS.DAT" AS "OUBTRANS.DAT"

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  NUBFile = FREEFILE
  OPEN "nUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBCust) / UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBCust, Cnt&, UBCustRec(1)
    ThisTran& = UBCustRec(1).LastTrans
    DO WHILE ThisTran& > 0
      GET UBFile, ThisTran&, UBTranRec(1)
      SELECT CASE UBTranRec(1).TransType
      'CASE TranAppliedDeposit, TranAppliedDeposit + 100
      '  IF UBTranRec(1).TransAmt = 0 THEN
      '    ZApp = ZApp + 1
      '  END IF
      '  TDepApp# = uRound#(TDepApp# + UBTranRec(1).TransAmt)
      '  App = App + 1
      CASE TranDepositPayment, TranDepositPayment + 100
        IF UBTranRec(1).TransAmt = 0 THEN
          ZAmt = ZAmt + 1
        END IF
        'TDepAmt# = uRound#(TDepAmt# + UBTranRec(1).TransAmt)
        'Amt = Amt + 1
      'CASE TranRefundDeposit, TranRefundDeposit + 100
      '  IF UBTranRec(1).TransAmt = 0 THEN
      '    ZRef = ZRef + 1
      '  END IF
      '  TDepRef# = uRound#(TDepRef# + UBTranRec(1).TransAmt)
      '  Ref = Ref + 1
      END SELECT
      ThisTran& = UBTranRec(1).PrevTrans
    LOOP

    'REDIM UBTranRec(1) AS UBTransRecType
    'TBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    'RevBal# = 0
    'FOR RevCnt = 1 TO 15
    '  RevBal# = Round#(RevBal# + UBCustRec(1).CurrRevAmts(RevCnt))
    'NEXT
    'IF RevBal# <> TBal# THEN
    '  UBCustRec(1).CurrBalance = RevBal#
    '  UBCustRec(1).PrevBalance = 0
    '  'PUT UBCust, Cnt&, UBCustRec(1)
    '  STOP
    '  PRINT Cnt&, QPTrim$(UBCustRec(1).CustName); TAB(40); RevBal#; TAB(50); TBal#; TAB(60); UBCustRec(1).Status
    'END IF
  NEXT
  CLOSE


LOCATE 9, 1
PRINT "Finished."
END

    IF RevBal# <> TBal# THEN
      'STOP
      LPRINT Cnt&, QPTrim$(UBCustRec(1).CustName); TAB(40); RevBal#; TAB(50); TBal#; TAB(60); UBCustRec(1).Status
    END IF

