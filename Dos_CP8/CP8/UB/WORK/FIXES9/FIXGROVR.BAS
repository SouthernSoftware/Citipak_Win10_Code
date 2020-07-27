DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.bi'
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))

  CLS
  PRINT
  PRINT "Correcting OLD Citipak converted Deposits. . ."
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  NumOFTR& = LOF(UBTrans) \ UBTransRecLen
  FOR Trans& = 1 TO NumOFTR&
    LOCATE 6, 1: PRINT "Processing:"; Trans&; "of"; NumOFTR&;
    GET UBTrans, Trans&, UBTrans(1)
    IF UBTrans(1).TransType = 106 THEN
      IF uRound#(UBTrans(1).TransAmt) > 0 THEN
        GET UBCust, UBTrans(1).CustAcctNo, UBCustRec(1)
        IF UBCustRec(1).DepositAmt < uRound#(UBTrans(1).TransAmt) THEN
          UBCustRec(1).DepositAmt = uRound#(UBTrans(1).TransAmt)
          PUT UBCust, UBTrans(1).CustAcctNo, UBCustRec(1)
        END IF
        Fixed = Fixed + 1
        UBTrans(1).TransType = 7
        LSET UBTrans(1).TransDesc = QPTrim$(UBTrans(1).TransDesc) + " CFO"
        UBTrans(1).RevAmt(1) = UBTrans(1).TransAmt
      ELSE
        UBTrans(1).CustAcctNo = 0
      END IF
      PUT UBTrans, Trans&, UBTrans(1)
    END IF
  NEXT
  CLOSE
  SHELL "UBMISC RELINK"
  CLS
  PRINT
  PRINT "Correction Complete. "
  PRINT
  PRINT "Changed:"; Fixed
  CALL KillFile("FIXGROVR.EXE")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

