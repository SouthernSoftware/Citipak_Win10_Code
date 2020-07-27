DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
  PRINT "Invalid Estimates. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))

  PRINT
  PRINT
  LPRINT "Zero Estimates"
  LPRINT
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidOne = False
    IF INSTR(UBCustRec(1).CustType, "MON") THEN
    'FOR MCnt = 1 TO 7
    '  IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
    '    WhatAmt& = UBCustRec(1).LocMeters(MCnt).AvgUse
    '    IF WhatAmt& <= 2 THEN
    '      DidOne = True
    '    END IF
    '  END IF
    'NEXT

    'IF DidOne THEN
      DidCnt = DidCnt + 1
     ' LPRINT UBCustRec(1).CustName, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb, Cnt&
     ' LineCnt = LineCnt + 1
     ' IF LineCnt > 60 THEN
     '   LPRINT CHR$(12)
     '   LineCnt = 1
     ' END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Zero Estimates:"; DidCnt

