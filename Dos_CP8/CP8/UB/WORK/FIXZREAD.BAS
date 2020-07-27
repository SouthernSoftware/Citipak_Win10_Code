DECLARE SUB SmallPause ()
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumb%)

  '$INCLUDE: 'pulacust.BI'


  CLS
  PRINT
  PRINT
  'PRINT "Processing Blank Readings. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFileBak = FREEFILE
  OPEN "UBCUST.bak" FOR RANDOM SHARED AS UBFileBak LEN = UBCustRecLen

  UBFileFix = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFileFix LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFileFix) \ UBCustRecLen

  GET UBFileBak, 26173, UBCustRec(1)
  PUT UBFileFix, 26173, UBCustRec(1)

  GET UBFileBak, 7527, UBCustRec(1)
  PUT UBFileFix, 7527, UBCustRec(1)

  GET UBFileBak, 7528, UBCustRec(1)
  PUT UBFileFix, 7528, UBCustRec(1)

  GET UBFileBak, 10824, UBCustRec(1)
  PUT UBFileFix, 10824, UBCustRec(1)

  GET UBFileBak, 19321, UBCustRec(1)
  PUT UBFileFix, 19321, UBCustRec(1)

  GET UBFileBak, 3924, UBCustRec(1)
  PUT UBFileFix, 3924, UBCustRec(1)



 ' FOR Cnt& = 1 TO NumOfRecs&
 '   LOCATE , 1
 '   PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
 '   GET UBFile, Cnt&, UBCustRec(1)
 '   DidEm = 0
 '   IF UBCustRec(1).LocMeters(1).CurRead < 0 THEN
 '     UBCustRec(1).LocMeters(1).CurRead = 0
 '     DidEm = -1
 '   END IF
 '   IF UBCustRec(1).LocMeters(1).PrevRead < 0 THEN
 '     UBCustRec(1).LocMeters(1).PrevRead = 0
 '     DidEm = -1
 '   END IF
 '
'      IF DidEm = -1 THEN
'        FixedCnt = FixedCnt + 1
'        PUT UBFile, Cnt&, UBCustRec(1)
'      END IF
'  '  SmallPause
'  NEXT



  CLOSE

  PRINT
  PRINT
  PRINT
  PRINT "Done"; FixedCnt
  'CALL KillFile("fixZRead.exe")

SUB SmallPause
ST! = TIMER
st2! = ST! + .008
DO
LOOP UNTIL TIMER > st2!
END SUB

