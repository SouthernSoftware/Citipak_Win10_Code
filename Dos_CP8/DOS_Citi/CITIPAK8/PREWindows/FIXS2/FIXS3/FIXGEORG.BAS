DECLARE FUNCTION uRound# (N#)
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num% (TheDate$)
  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.bi'

  TRDate = Date2Num("02-17-1999")

  CLS
  PRINT
  PRINT
  PRINT "Correcting Balances...."
  REDIM UBTrans(1)        AS UBTransRecType
  UBTranRecLen = LEN(UBTrans(1))
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  TRFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS TRFile LEN = UBTranRecLen
  
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBFile) / UBCustRecLen
  FOR Cnt& = 1 TO NumOfCust&
    GET UBFile, Cnt&, UBCustRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfCust&;
    LastTran& = UBCustRec(1).LastTrans
    IF LastTran& > 0 THEN
      GET TRFile, LastTran&, UBTrans(1)
      IF UBTrans(1).TransType = TranPenaltyCharge THEN
        IF UBTrans(1).TransDate = TRDate THEN
          UBCustRec(1).CurrBalance = uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
          UBCustRec(1).PrevBalance = 0
          PUT UBFile, Cnt&, UBCustRec(1)
          Found = Found + 1
        END IF
      END IF
    END IF
  NEXT
  CLOSE

PRINT
PRINT
PRINT
PRINT "Corrected:"; Found

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

