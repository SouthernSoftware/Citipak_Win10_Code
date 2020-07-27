DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumb%)

  '$INCLUDE: 'pubcust.BI'


  CLS
  PRINT
  PRINT
  PRINT "Clearing Customer Type Field. . ."
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
    UBCustRec(1).CustType = ""
    FixedCnt = FixedCnt + 1
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; FixedCnt
  CALL KillFile("fixctype.exe")

SUB SmallPause
st! = TIMER
st2! = st! + .008
DO
LOOP UNTIL TIMER > st2!
END SUB

