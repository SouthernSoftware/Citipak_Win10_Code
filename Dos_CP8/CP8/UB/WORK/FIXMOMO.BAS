DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)


  '$INCLUDE: 'pulacust.BI'


  CLS
  PRINT
  PRINT
  PRINT "Processing Damaged Accounts. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFileBak = FREEFILE
  OPEN "UBCUST.bak" FOR RANDOM SHARED AS UBFileBak LEN = UBCustRecLen

  UBFileFix = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFileFix LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFileFix) \ UBCustRecLen

  GET UBFileBak, 7685, UBCustRec(1)
  PUT UBFileFix, 7685, UBCustRec(1)

  GET UBFileBak, 28581, UBCustRec(1)
  PUT UBFileFix, 28581, UBCustRec(1)


  CLOSE

  PRINT
  PRINT
  PRINT
  PRINT "Done.";

  CALL KillFile("fixmomo.exe")

