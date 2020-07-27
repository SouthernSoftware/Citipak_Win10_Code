DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Books 11, 12, 13, 14, 15, 18. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    SELECT CASE UBCustRec(1).Book
    CASE "11", "12", "13", "14", "15", "18"
      Serv5$ = QPTrim$(UBCustRec(1).Serv(5).RateCode)
      IF LEN(Serv5$) > 0 THEN
        FOR MtrCnt = 1 TO 7
          IF UBCustRec(1).LocMeters(MtrCnt).MTRType = "W" THEN
            UBCustRec(1).LocMeters(MtrCnt).CurRead = UBCustRec(1).LocMeters(MtrCnt).PrevRead
          END IF
        NEXT
        PUT UBCFile, Cnt&, UBCustRec(1)
      END IF
    CASE ELSE
    END SELECT

  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIXIRRIG.EXE"

