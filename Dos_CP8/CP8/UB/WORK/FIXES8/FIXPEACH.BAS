DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Book 01. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          IF UBCustRec(1).LocMeters(MtrCnt).AvgUse > 0 THEN
            UBCustRec(1).LocMeters(MtrCnt).CurRead = UBCustRec(1).LocMeters(MtrCnt).PrevRead + UBCustRec(1).LocMeters(MtrCnt).AvgUse
          END IF
        END IF
      NEXT
      PUT UBCFile, Cnt&, UBCustRec(1)

  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIXPEACH.EXE"

