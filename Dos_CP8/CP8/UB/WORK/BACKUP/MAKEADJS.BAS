DEFINT A-Z
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION uRound# (N#)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.bi'

   CLS
   PRINT "Creating BAD Debt Transactions. . ."

   REDIM UBCustRec(1)   AS NewUBCustRecType
   UBCustRecLen = LEN(UBCustRec(1))
   REDIM UBTransRec(1) AS UBTransRecType
   UBTransRecLen = LEN(UBTransRec(1))

   TransDate = Date2Num(DATE$)

   UBCust = FREEFILE
   OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

   NumOfCust& = LOF(UBCust) / UBCustRecLen

   FOR zz& = 1 TO NumOfCust&
     LOCATE 5, 1
     PRINT "Processing:"; zz&; " of"; NumOfCust&;
     GET UBCust, zz&, UBCustRec(1)
     IF UBCustRec(1).Book = "50" OR UBCustRec(1).Book = "60" THEN
       RealBalance# = uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
       IF RealBalance# > 0 THEN
         CustAcct& = zz&
         GOSUB MakeTrans
         DidCnt = DidCnt + 1
       END IF
     END IF
   NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Created:"; DidCnt
  KILL "MAKEADJS.EXE"
  END




MakeTrans:

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTransRecLen

  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  UBTransRec(1).TransDate = TransDate

  TransAmt# = 0
  FOR Cnt = 1 TO 15
    TransAmt# = uRound#(TransAmt# + -UBCustRec(1).CurrRevAmts(Cnt))
    UBTransRec(1).RevAmt(Cnt) = -UBCustRec(1).CurrRevAmts(Cnt)
  NEXT
  
  UBTransRec(1).CustStatus = UBCustRec(1).STATUS
  UBTransRec(1).CustAcctNo = CustAcct&
  UBTransRec(1).BillMsg = "BAD DEBT WRITE OFF"
  UBTransRec(1).TransAmt = TransAmt#
  
  UBTransRec(1).TransDesc = "BAD DEBT ADJUSTMENT"

  FOR RevCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(RevCnt) = 0
    UBCustRec(1).CurrBalance = 0
  NEXT
  UBTransRec(1).TransType = TranDownwardAdjustment

  UBTransRec(1).RunBalance = 0  'Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)


  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans

  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&
  
  PUT UBCust, UBTransRec(1).CustAcctNo, UBCustRec(1)
  PUT UBTran, NextTranRecs&, UBTransRec(1)
  CLOSE UBTran

RETURN

