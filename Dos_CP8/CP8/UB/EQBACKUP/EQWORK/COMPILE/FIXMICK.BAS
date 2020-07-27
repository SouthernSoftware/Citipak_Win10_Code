DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

  '$INCLUDE: 'newcust.BI'

  CLS
  DIM DidEm AS INTEGER

  BDate = Date2Num("12-31-1979")
  GDate = Date2Num("08-14-2000")

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
    PRINT "Processing Customer:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = 0
    IF UBCustRec(1).Book = "03" THEN
      DidEm = -1
      
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).PrevRead > 0 THEN
          UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).PrevRead
        END IF
      NEXT
      FixedCnt = FixedCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF

SkipThisOne:
  NEXT
  CLOSE UBFile

  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; FixedCnt
  KILL "FIXMICK.EXE"


