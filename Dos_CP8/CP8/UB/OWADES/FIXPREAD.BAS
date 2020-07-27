DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
  PRINT "Adjusting Est. Flags. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))

  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidOne = False
    SELECT CASE UBCustRec(1).Book
    'CASE "01"
    '  DidOne = True
    '  FOR MCnt = 1 TO 7
    '    'WhatAmt& = UBCustRec(1).LocMeters(MCnt).AvgUse
    '    'IF WhatAmt& > 0 THEN
    '    '  UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead - WhatAmt&
    '    '  UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead - WhatAmt&
    '    'END IF
    '    'UBCustRec(1).LocMeters(MCnt).ReadFlag = "E"
    '  NEXT
    CASE "03", "04"

    CASE ELSE
      DidOne = True
      FOR MCnt = 1 TO 7
        UBCustRec(1).LocMeters(MCnt).ReadFlag = "E"
      NEXT
    END SELECT
    IF DidOne THEN
      DidCnt = DidCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished:"

