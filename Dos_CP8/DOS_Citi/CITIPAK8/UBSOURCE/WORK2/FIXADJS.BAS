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

  PDate = Date2Num("07-08-1998")

  CLS
  PRINT "DO NOTHING UNTIL THIS PROGRAM HAS COMPLETED"
  PRINT
  PRINT
  PRINT "Processing Transactions. . ."
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
    IF UBTrans(1).TransType = TranUpwardAdjustment THEN
      IF UBTrans(1).TransDate <> PDate THEN
        IF UBTrans(1).CustAcctNo < 0 THEN
          AcctNo& = ABS(UBTrans(1).CustAcctNo)
          GET UBCust, AcctNo&, UBCust(1)
          CurBal# = 0
          FOR zCnt = 1 TO 15
            CurBal# = uRound#(CurBal# + UBTrans(1).RevAmt(zCnt))
            UBCust(1).CurrRevAmts(zCnt) = uRound#(UBCust(1).CurrRevAmts(zCnt) + UBTrans(1).RevAmt(zCnt))
          NEXT
          UBCust(1).CurrBalance = CurBal#
          UBCust(1).PrevBalance = 0
          PUT UBCust, AcctNo&, UBCust(1)
          UBTrans(1).CustAcctNo = ABS(UBTrans(1).CustAcctNo)
          'UBTrans(1).PrevTrans = 0
          'Fixed& = Fixed& + 1
          PUT UBFile, Cnt&, UBTrans(1)
        END IF
      END IF
    END IF
  NEXT
  CLOSE
SHELL "ubmisc relink"
PRINT
PRINT
PRINT
PRINT "Processing complete"
PRINT
PRINT

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

