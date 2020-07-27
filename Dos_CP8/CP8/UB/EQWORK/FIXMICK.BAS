DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

  '$INCLUDE: 'newcust.BI'

  CLS
  DIM DidEm AS INTEGER

  BDate = Date2Num("04-14-2010")
  GDate = Date2Num("07-14-2010")

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
 
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).CurDate = BDate THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = GDate
        DidEm = -1
      END IF
    NEXT
    IF DidEm = -1 THEN
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


