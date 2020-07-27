DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))
  CLS
  LOCATE 2, 1, 0

  PRINT "Adjusting Touch Readings. . ."
  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      DidOne = 0
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).MTRType = "T" THEN
          DidOne = -1
          UBCustRec(1).LocMeters(MCnt).MtrMulti = 1000
          Prev# = UBCustRec(1).LocMeters(MCnt).PrevRead
          Curr# = UBCustRec(1).LocMeters(MCnt).CurRead
          IF Prev# >= 1000 THEN
            Prev# = Prev# / 1000
            UBCustRec(1).LocMeters(MCnt).PrevRead = Prev#
          END IF
          IF Curr# >= 1000 THEN
            Curr# = Curr# / 1000
            UBCustRec(1).LocMeters(MCnt).CurRead = Curr#
          END IF
        END IF
      NEXT
      IF DidOne THEN
        DidCnt = DidCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Processing Complete."
  PRINT
  PRINT "Corrected:"; DidCnt;
  'KILL "FIXTOUCH.exe"

