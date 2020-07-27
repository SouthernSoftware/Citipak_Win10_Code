DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE SUB killfile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CLS

  PRINT
  PRINT
  PRINT "Changing Zip Codes. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR zz& = 1 TO NumOfRecs&
    GET UBFile, zz&, UBCustRec(1)
    SmallPause
    LOCATE 7, 1
    PRINT "   Zipcode: "; UBCustRec(1).zipcode;
    Zip$ = QPTrim$(UBCustRec(1).zipcode)
    LOCATE 10, 1
    PRINT "Processing:"; zz&; "of"; NumOfRecs&;
    IF RIGHT$(Zip$, 4) <> "9999" THEN
      ZipCnt = ZipCnt + 1
      Zip$ = LEFT$(UBCustRec(1).zipcode, 5)
      LSET UBCustRec(1).zipcode = Zip$
      PUT UBFile, zz&, UBCustRec(1)
    END IF
  NEXT


  LOCATE 15, 1
  PRINT "Done: "; ZipCnt

  CLOSE
  CALL killfile("fixzcode.exe")

SUB SmallPause
  st1! = TIMER
  st2! = st1! + .22
  DO
  LOOP UNTIL TIMER > st2!

END SUB

