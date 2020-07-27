DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION Round# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TranDate1 = Date2Num("10-31-1997")

  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."
 
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  NUBFile = FREEFILE
  OPEN "NUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen
  
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    'IF UBTranRec(1).CustAcctNo = 1416 THEN
      IF UBTranRec(1).TransDate >= TranDate1 THEN
        Removed = Removed + 1
      ELSE
        PUT NUBFile, , UBTranRec(1)
      END IF
    'ELSE
    '  PUT NUBFile, , UBTranRec(1)
    'END IF
  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Removed

