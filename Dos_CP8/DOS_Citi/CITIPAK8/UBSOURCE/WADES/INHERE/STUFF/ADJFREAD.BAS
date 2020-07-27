DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
'  PRINT "Adjusting Previous Readings. . ."

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
    SELECT CASE Cnt&
    CASE 955, 1852, 1919
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y" THEN
          DidOne = True
          WhatAmt& = (UBCustRec(1).LocMeters(MCnt).AvgUse * 2)
          UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead - WhatAmt&
        END IF
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)

    END SELECT
  NEXT
  CLOSE UBFile

