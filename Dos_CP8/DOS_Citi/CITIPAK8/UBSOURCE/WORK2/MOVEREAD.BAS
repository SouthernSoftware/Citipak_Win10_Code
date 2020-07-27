DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  RDate = Date2Num("01-09-1998")

  CLS
  PRINT
  PRINT
  PRINT "Processing Water Readings for Books 1, 2, & 3. . ."
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
    IF UBCustRec(1).Status = "A" THEN
      'IF LEFT$(UBCustRec(1).Book, 1) = "0" THEN STOP
      Book = VAL(UBCustRec(1).Book)
      SELECT CASE Book
      CASE 1, 2, 3
        FOR MCnt = 1 TO 7
          IF UBCustRec(1).LocMeters(MCnt).MtrType = "C" OR UBCustRec(1).LocMeters(MCnt).MtrType = "W" THEN
            UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBCustRec(1).LocMeters(MCnt).CurDate = RDate
          END IF
        NEXT
        PUT UBFile, Cnt&, UBCustRec(1)
      END SELECT

      
    END IF
  NEXT
  CLOSE

PRINT
PRINT
PRINT
PRINT "Done."

