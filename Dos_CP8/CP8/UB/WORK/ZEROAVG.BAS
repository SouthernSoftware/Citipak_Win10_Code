DEFINT A-Z
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Zeroing Average Consumption. . ."
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
    IF UBCustRec(1).DelFlag <> 0 THEN
      deleted = deleted + 1
    END IF

'    FOR RCnt = 1 TO 7
'      UBCustRec(1).LocMeters(RCnt).AvgUse = 0
'    NEXT
'    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "deleted:"; deleted

CALL killfile("zeroavg.exe")

