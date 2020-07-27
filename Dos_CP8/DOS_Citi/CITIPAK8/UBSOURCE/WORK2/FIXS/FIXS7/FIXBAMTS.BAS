DEFINT A-Z
DECLARE FUNCTION Date2Num% (Today$)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'cmfiles.BI'

'  TDate = Date2Num%("02-17-1998")

  CLS
  PRINT "Correcting UB Invalid Check Amount. . ."
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM CMTranRec(1) AS CMTransRecType

  UBTranRecLen = LEN(UBTranRec(1))
  CMTranRecLen = LEN(CMTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&

  GET UBFile, Cnt&, UBTranRec(1)
  LOCATE 3, 1: PRINT "Scanning:"; Cnt&; " of"; TNumOfRecs&
  IF UBTranRec(1).CheckAmount < -100000 THEN
    UBTranRec(1).CheckAmount = 0
    Fixed = Fixed + 1
    PUT UBFile, Cnt&, UBTranRec(1)
  END IF
  NEXT
  CLOSE

  LOCATE 5, 1
  PRINT "Corrected:"; Fixed
  Fixed = 0

  LOCATE 10, 1
  PRINT "Correcting CM Invalid Check Amount. . ."
  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTranRecLen
  TNumOfRecs& = LOF(CMFile) / CMTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&

  GET CMFile, Cnt&, CMTranRec(1)
  LOCATE 13, 1: PRINT "Scanning:"; Cnt&; " of"; TNumOfRecs&
  IF CMTranRec(1).TransCheck < -100000 THEN
    CMTranRec(1).TransCheck = 0
    Fixed = Fixed + 1
    PUT UBFile, Cnt&, CMTranRec(1)
  END IF
  NEXT

  CLOSE

  LOCATE 15, 1
  PRINT "Corrected:"; Fixed

