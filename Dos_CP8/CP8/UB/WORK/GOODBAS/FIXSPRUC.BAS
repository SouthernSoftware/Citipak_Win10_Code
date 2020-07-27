DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Previous Readings. . ."
  REDIM UBCustRec(1) AS NewUBCustRecTypeW
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
    DidEm = 0
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).PrevRead > 0 THEN
        UBCustRec(1).LocMeters(MCnt).PrevRead = (UBCustRec(1).LocMeters(MCnt).PrevRead * .001)
        UBCustRec(1).LocMeters(MCnt).MtrMulti = 1000
        DidEm = -1
      END IF
    NEXT

    IF DidEm THEN
      FixedCnt = FixedCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; FixedCnt
  CALL KillFile("fixspruc.exe")

