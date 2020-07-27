DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    GET UBCFile, Cnt&, UBCustRec(1)
    FOR FRCnt = 1 TO 4
      IF UBCustRec(1).FlatRates(FRCnt).REVSRC = 3 THEN
        SELECT CASE UBCustRec(1).FlatRates(FRCnt).FRAMT
        CASE 8.28#
          UBCustRec(1).FlatRates(FRCnt).FRAMT = 9
          Cnt8 = Cnt8 + 1
        CASE 9.28#
          UBCustRec(1).FlatRates(FRCnt).FRAMT = 10
          Cnt9 = Cnt9 + 1
        CASE ELSE
          NoFix = NoFix + 1
        END SELECT
      END IF
    NEXT
    PUT UBCFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  PRINT
  PRINT "8.28 Changed:"; Cnt8
  PRINT "9.28 Changed:"; Cnt9
  PRINT "   Unchanged:"; NoFix
  KillFile "FIXBEECH.EXE"


