DEFINT A-Z
  
  '$INCLUDE: 'ubtrans.BI'


  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Deposit Information. . ."

  PRINT
  PRINT

  UBTFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen

  GET UBTFile, 19114, UBTranRec(1)
  UBTranRec(1).RevAmt(1) = -40
  PUT UBTFile, 19114, UBTranRec(1)
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."
  KILL "FIXMEBAN.EXE"

