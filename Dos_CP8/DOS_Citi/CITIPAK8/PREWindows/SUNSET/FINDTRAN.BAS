DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  GoodDate = Date2Num("09-21-1999")
  'BadDate = Date2Num("09-21-1999")


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Deposit Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    IF UBTranRec(1).TransDate = GoodDate THEN
      SELECT CASE UBTranRec(1).TransType
      CASE TranPenaltyCharge
        GET #UBCFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        UBCustRec(1).CurrRevAmts(4) = uRound(UBCustRec(1).CurrRevAmts(4) - UBTranRec(1).RevAmt(4))
        UBCustRec(1).CurrBalance = uRound(UBCustRec(1).CurrBalance - UBTranRec(1).RevAmt(4))
        PUT #UBCFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        PUT UBTFile, Cnt&, UBTranRec(2)
        FoundIt = FoundIt + 1
      END SELECT
    END IF
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; FoundIt

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

