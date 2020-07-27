DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION uRound# (DblNum#)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))
  REDIM UBCust(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCust(1))

  PDate = Date2Num("07-15-1998")

  CLS
  PRINT "DO NOTHING UNTIL THIS PROGRAM HAS COMPLETED"
  PRINT
  PRINT
  PRINT "Removing Penalty Transactions  07-15-98. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBFile LEN = UBTransRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBTransRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBTrans(1)
    IF UBTrans(1).TransType = TranPenaltyCharge THEN
      IF UBTrans(1).TransDate = PDate THEN
        GET UBCust, UBTrans(1).CustAcctNo, UBCust(1)
        'LPRINT USING "#######"; UBTrans(1).CustAcctNo
        Penalty# = UBTrans(1).RevAmt(2)
        UBCust(1).CurrRevAmts(2) = uRound#(UBCust(1).CurrRevAmts(2) - Penalty#)
        UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - Penalty#)
        PUT UBCust, UBTrans(1).CustAcctNo, UBCust(1)
        UBTrans(1).CustAcctNo = -UBTrans(1).CustAcctNo
        PUT UBFile, Cnt&, UBTrans(1)
        Fixed& = Fixed& + 1
      END IF
    END IF
  NEXT
  CLOSE
'SHELL "ubmisc relink"
PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed&
PRINT
PRINT

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

