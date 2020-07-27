DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION Round# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TranDate1 = Date2Num("07-02-1997")
  'TranDate2 = Date2Num("03-11-1997")
  'PastDate = Date2Num("11-15-1996")

  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  NUBFile = FREEFILE
  OPEN "NUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen
  
  CUBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CUBFile LEN = UBCustRecLen

  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransDate = TranDate1 THEN
      IF UBTranRec(1).TransType = TranPenaltyCharge THEN
        Removed = Removed + 1
      ELSE
        PUT NUBFile, , UBTranRec(1)
      END IF
    ELSE
      PUT NUBFile, , UBTranRec(1)
    END IF

  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Removed

'      IF UBTranRec(1).TransAmt > 0 THEN
'        PUT NUBFile, , UBTranRec(1)
'      ELSE
'        Removed = Removed + 1
'      END IF

