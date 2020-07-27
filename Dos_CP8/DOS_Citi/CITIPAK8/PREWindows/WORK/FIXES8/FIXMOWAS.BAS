DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Books 01, 02, 03. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    SELECT CASE UBCustRec(1).Book
    CASE "01", "02", "03"
      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).CurRead > 0 THEN
          CurRead$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt).CurRead))
          ReadLen = LEN(CurRead$)
          IF ReadLen > 1 THEN
            CurRead$ = LEFT$(CurRead$, ReadLen - 1) + "0"
            UBCustRec(1).LocMeters(MtrCnt).CurRead = VAL(CurRead$)
          END IF
        END IF
      NEXT
      PUT UBCFile, Cnt&, UBCustRec(1)
    CASE ELSE
    END SELECT

  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIXMOWAS.EXE"

