DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (FileName$)
DECLARE FUNCTION QPValL& (Numb$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Current Readings. . ."

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
    DidOne = 0
    FOR MTCnt = 1 TO 7
      ThisRead$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MTCnt).CurRead))
      ThisLen = LEN(ThisRead$)
      IF ThisLen > 3 THEN
        IF RIGHT$(ThisRead$, 3) = "000" THEN
          Good$ = LEFT$(ThisRead$, ThisLen - 3)
          CurRead& = QPValL(Good$)
          UBCustRec(1).LocMeters(MTCnt).CurRead = CurRead&
          DidOne = -1
        END IF
      END IF
    NEXT
    IF DidOne THEN
      Fixed = Fixed + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed

