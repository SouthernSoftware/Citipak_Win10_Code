DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))
  CLS
  LOCATE 2, 1, 0
  PRINT "Adjusting Current Readings. . ."
  
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
        IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 AND UBCustRec(1).LocMeters(MCnt).PrevRead >= 0 THEN
          DidOne = -1
          CurrRead# = UBCustRec(1).LocMeters(MCnt).CurRead
          PrevRead# = UBCustRec(1).LocMeters(MCnt).PrevRead
          MeterConsp& = CurrRead# - PrevRead#
          IF MeterConsp& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(PrevRead#)) - 1)
            MeterConsp& = (MaxMeterAmt& - PrevRead#) + CurrRead#
          END IF
          UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead + MeterConsp&
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
  PRINT "Adjusted:"; DidCnt; "Readings."
  KILL "setreads.exe"

