DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

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
    FOR ZZ = 1 TO 7
      Curr& = UBCustRec(1).locmeters(ZZ).curread
      Prev& = UBCustRec(1).locmeters(ZZ).PrevRead
      IF Curr& > 0 THEN
        IF Prev& > Curr& THEN
          Prevr$ = QPTrim$(STR$(Prev&))
          Prevr$ = LEFT$(Prevr$, (LEN(Prevr$) - 1))
          UBCustRec(1).locmeters(ZZ).PrevRead = VAL(Prevr$)
          UBCustRec(1).locmeters(ZZ).MTRMulti = 10
          Cnt10 = Cnt10 + 1
          PUT UBCFile, Cnt&, UBCustRec(1)
        END IF
      END IF
    NEXT
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  PRINT "Multi=10:"; Cnt10
  PRINT "Done"
  KillFile "FIXMULTI.EXE"


