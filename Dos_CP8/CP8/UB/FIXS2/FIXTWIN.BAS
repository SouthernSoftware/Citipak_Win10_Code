DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting FIRE Revenue & Coping Postal Route Codes. . ."

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
    IF UBCustRec(1).DelFlag = 0 THEN
      UBCustRec(1).Serv(5).RateCode = "FIRE"
      PostRT$ = QPTrim$(UBCustRec(1).ServAddr)
      IF LEN(PostRT$) = 4 THEN
        UBCustRec(1).PostRte = PostRT$
      ELSE
        UBCustRec(1).PostRte = ""
      END IF
      FOR zz = 1 TO 4
        Char$ = MID$(UBCustRec(1).PostRte, zz, 1)
        IF Char$ = "O" THEN
          MID$(UBCustRec(1).PostRte, zz, 1) = "0"
        END IF
      NEXT
      Fixed = Fixed + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed

