DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumb%)

  '$INCLUDE: 'newcust.BI'

  RDate1 = Date2Num("01-01-1903")
  RDate2 = Date2Num("12-31-1903")

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
    DidEm = 0
    FOR MCnt = 1 TO 7
      SELECT CASE UBCustRec(1).LocMeters(MCnt).CurDate
      CASE RDate1 TO RDate2
        WhatDate$ = Num2Date(UBCustRec(1).LocMeters(MCnt).CurDate)
        WhatDate$ = LEFT$(WhatDate$, 6) + "2003"
        UBCustRec(1).LocMeters(MCnt).CurDate = Date2Num(WhatDate$)
        DidEm = -1
      END SELECT
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
  CALL KillFile("fixrdate.exe")

