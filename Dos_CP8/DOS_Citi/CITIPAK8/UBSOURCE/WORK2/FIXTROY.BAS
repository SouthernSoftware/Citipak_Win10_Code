DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  PRINT
  PRINT
  PRINT "Processing Readings. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidOne = 0
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).MtrType <> " " THEN
        IF UBCustRec(1).LocMeters(MCnt).MtrType <> "T" THEN
          UBCustRec(1).LocMeters(MCnt).MtrMulti = 1000
          Pre$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead))
          Pre$ = LEFT$(Pre$, (LEN(Pre$) - 1))
          UBCustRec(1).LocMeters(MCnt).PrevRead = VAL(Pre$)
          DidOne = -1
        ELSE
          NotOne = NotOne + 1
        END IF
      END IF
    NEXT
    IF DidOne THEN
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF

  NEXT
  CLOSE
PRINT
PRINT
PRINT
PRINT "Done."

