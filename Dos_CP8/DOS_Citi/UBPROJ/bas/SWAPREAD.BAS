DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  CONST False = 0, True = NOT False

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  LOCATE 2, 1, 0
  PRINT "Processing Book 7. . ."

  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Book = "07" THEN
      FOR Mtr = 1 TO 7
        SWAP UBCustRec(1).LocMeters(Mtr).PrevRead, UBCustRec(1).LocMeters(Mtr).CurRead
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Corrected:"; Fixed
  PRINT
  PRINT

