DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  CONST False = 0, True = NOT False

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  LOCATE 2, 1, 0
  PRINT "Coping Readings. . ."

  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    FOR Mtr = 1 TO 7
      UBCustRec(1).LocMeters(Mtr).PrevRead = UBCustRec(1).LocMeters(Mtr).CurRead
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 10, 1
  PRINT "Processing Complete."
  PRINT
  PRINT

