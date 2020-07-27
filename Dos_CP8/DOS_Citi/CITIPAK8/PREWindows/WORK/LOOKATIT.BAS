DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  Date1 = Date2Num%("08-15-2002")
  
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  CustLen = LEN(UBCust(1))

  'KillFile "UBTRANS.bak"
  'NAME "UBTRANS.DAT" AS "UBTRANS.bak"

  'UBFile1 = FREEFILE
  'OPEN "UBTRANS.bak" FOR RANDOM SHARED AS UBFile1 LEN = UBTranRecLen

  UBFile2 = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile2 LEN = UBTranRecLen

  UBFile3 = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile3 LEN = CustLen

  'TNumOfRecs& = LOF(UBFile1) / UBTranRecLen
  CNumOfRecs& = LOF(UBFile3) / CustLen

  'FOR Cnt& = 1 TO TNumOfRecs&
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; CNumOfRecs&;
    'GET UBFile1, Cnt&, UBTranRec(1)
    GET UBFile3, Cnt&, UBCust(1)
    'IF UBTranRec(1).TransDate = Date1 THEN
     ' SELECT CASE UBTranRec(1).TransType
     ' CASE TranUtilityBill
     '   GET UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)

        SELECT CASE UBCust(1).BILLCYCL
        CASE 11, 12, 14
          TAmt# = 0
          FOR RCnt = 1 TO 15
            TAmt# = uRound(TAmt# + UBCust(1).CurrRevAmts(RCnt))
            'UBCust(1).CurrRevAmts(1) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
          NEXT
          IF TAmt# <> uRound#(UBCust(1).CurrBalance + UBCust(1).PrevBalance) THEN
            IF UBCust(1).LastTrans > 0 THEN
              GET #UBFile2, UBCust(1).LastTrans, UBTranRec(1)
              SELECT CASE UBTranRec(1).TransType
              CASE 1     'bill
                Bill = Bill + 1
              CASE 4     'stdpayment
                Pay = Pay + 1
                'STOP
                IF UBTranRec(1).OperatorNumber <> 1 THEN
                '  Othr = Othr + 1
                '  STOP
                END IF
              CASE 6     'penalty
                Pnl = Pnl + 1
              CASE 12    'adjustment
                Adj = Adj + 1
              CASE ELSE
                STOP
              END SELECT
            END IF
          END IF
          'UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).TransAmt)
          'PUT UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
          'Removed = Removed + 1
        CASE ELSE
          'PUT UBFile2, , UBTranRec(1)
        END SELECT
      'CASE ELSE
      '  PUT UBFile2, , UBTranRec(1)
      'END SELECT
    'ELSE
    '  PUT UBFile2, , UBTranRec(1)
    'END IF
  NEXT

  CLOSE

LOCATE 9, 1
'PRINT " Copied:"; TTCnt&;
'PRINT "Removed:"; Removed

PRINT "Bill:"; Bill
PRINT " Pay:"; Pay, Othr
PRINT " Pen:"; Pnl
PRINT " Adj:"; Adj

END
KILL "UBTRANS.bak"
KILL "striptrn.exe"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

