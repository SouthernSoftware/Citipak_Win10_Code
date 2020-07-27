DEFINT A-Z
DECLARE FUNCTION uRound# (DblNum#)


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Billing Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN

 '     GOSUB AddRevs
      IF UBTranRec(1).RevAmt(8) > 0 THEN
        GET UBCFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        UBCustRec(1).CurrBalance = uRound(UBCustRec(1).CurrBalance - UBTranRec(1).RevAmt(8))
        UBCustRec(1).CurrRevAmts(8) = uRound(UBCustRec(1).CurrRevAmts(8) - UBTranRec(1).RevAmt(8))
        UBTranRec(1).TransAmt = uRound#(UBTranRec(1).TransAmt - UBTranRec(1).RevAmt(8))
        UBTranRec(1).RunBalance = uRound#(UBTranRec(1).RunBalance - UBTranRec(1).RevAmt(8))
        UBTranRec(1).RevAmt(8) = 0

        Fixed = Fixed + 1
        PUT UBCFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        PUT UBTFile, Cnt&, UBTranRec(1)
      END IF
    END IF
  NEXT

SkiptoNext:
  CLOSE

CALL killfile("fixeover.exe")

LOCATE 9, 1
PRINT "Fixed:"; Fixed

END
'AddRevs:
' RevTot# = 0
'
'
' NEXT
' RETURN

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .50000001#) / 100
END FUNCTION

