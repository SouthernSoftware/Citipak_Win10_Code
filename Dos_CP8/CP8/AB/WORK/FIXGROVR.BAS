DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT "Processing accounts. . ."
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  FOR zz& = 1 TO LOF(UBFile) / UBCustRecLen
    GET UBFile, zz&, UBCustRec(1)
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipEm
    END IF
    IF UBCustRec(1).TAXEXPT <> "N" THEN
      LPRINT zz&
      Bad = Bad + 1
    END IF
SkipEm:
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "bad:"; Bad
  CALL KillFile("fixgrovr.exe")

