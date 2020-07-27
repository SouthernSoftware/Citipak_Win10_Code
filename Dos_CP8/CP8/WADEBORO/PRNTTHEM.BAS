DEFINT A-Z
DECLARE FUNCTION qptrim$ (Text$)
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
  PRINT "Adjusting Previous Readings. . ."

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
    FOR MCnt = 1 TO 7
    IF (UBCustRec(1).Book = "03") OR (UBCustRec(1).Book = "04") OR (INSTR(UBCustRec(1).CustType, "MON") > 0) THEN
      HowMuch& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
      IF HowMuch& >= 50000 THEN
        DidOne = True
        EXIT FOR
      END IF
    END IF
    NEXT
    IF DidOne THEN
      DidCnt = DidCnt + 1
      FOR zz = 1 TO 15
        IF qptrim$(UBCustRec(1).Serv(zz).RateCode) = "WI" THEN
          LPRINT UBCustRec(1).CustName
          EXIT FOR
        END IF
      NEXT
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Adjusted:"; DidCnt

