DEFINT A-Z

  '$INCLUDE: 'ubtrans.BI'


  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction Information. . ."

  PRINT
  PRINT
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  GET UBTFile, 52063, UBTranRec(1)
  UBTranRec(1).RevAmt(1) = 60
  PUT UBTFile, 52063, UBTranRec(1)


SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Fixed: 1"
CALL killfile("fixface.exe")

