DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.BI'
  CONST True = -1, False = 0

  CLS
  PRINT
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))

  PRINT
  PRINT
  PRINT "Adjusting Estimate Flags"
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
    CASE "01", "03", "04", "07", "08", "09", "10"
'      IF INSTR(UBCustRec(1).CustType, "MON") <= 0 THEN
'        DidCnt = DidCnt + 1
'        FOR MCnt = 1 TO 7
'          IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
'            UBCustRec(1).LocMeters(MCnt).ReadFlag = "E"
'          END IF
'        NEXT
'        PUT UBFile, Cnt&, UBCustRec(1)
'      END IF
    CASE ELSE
      DidOne = 0
      FOR MCnt = 1 TO 7
        IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
          DidOne = -1
          UBCustRec(1).LocMeters(MCnt).ReadFlag = "Y"
        END IF
      NEXT
      IF DidOne THEN
        DidCnt = DidCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF

    END SELECT

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Adjusted:"; DidCnt

