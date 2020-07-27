DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (FileName$)
DECLARE FUNCTION QPValL& (Numb$)
DECLARE FUNCTION Date2Num (TheDate$)

  GoodDate = Date2Num("08-10-1998")

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Read Dates. . ."

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
    IF UBCustRec(1).Book = "03" THEN
      DidOne = 0
      FOR MTCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MTCnt).MtrType = "E" THEN
          UBCustRec(1).LocMeters(MTCnt).CurDate = GoodDate
          DidOne = -1
        END IF
      NEXT
      IF DidOne THEN
        Fixed = Fixed + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed

