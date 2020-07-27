DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  Date1 = Date2Num%("06-02-1999")

'  Date2 = Date2Num%("02-28-1997")
'  ActNo = 3822

  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  CustLen = LEN(UBCust(1))

  TranFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = UBTranRecLen
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = CustLen

  TNumOfRecs& = LOF(TranFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET TranFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate = Date1 THEN
      IF UBTranRec(1).CustAcctNo > 0 THEN
        GET CustFile, UBTranRec(1).CustAcctNo, UBCust(1)
        IF UBCust(1).Book = "50" THEN
          TBal# = 0
          FOR zz = 1 TO 15
            UBCust(1).CurrRevAmts(zz) = uRound(UBCust(1).CurrRevAmts(zz) - UBTranRec(1).RevAmt(zz))
            TBal# = uRound(TBal# + UBCust(1).CurrRevAmts(zz))
          NEXT
          UBCust(1).CurrBalance = TBal#
          UBCust(1).PrevBalance = 0
          PUT CustFile, UBTranRec(1).CustAcctNo, UBCust(1)
          Removed = Removed + 1
          UBTranRec(1).TransAmt = 0
          UBTranRec(1).CustAcctNo = 0
          PUT TranFile, Cnt&, UBTranRec(1)
        END IF
      END IF
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Removed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

