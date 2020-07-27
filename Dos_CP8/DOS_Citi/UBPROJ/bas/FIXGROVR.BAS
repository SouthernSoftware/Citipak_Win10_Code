DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT "Processing 'I' & 'B' status accounts. . ."
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Status = "I" OR UBCustRec(1).Status = "B" THEN
      UBCustRec(1).DepositAmt = 0
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing complete."
  CALL KillFile("fixgrovr.exe")

