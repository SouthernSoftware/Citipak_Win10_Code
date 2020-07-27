DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT "Processing. . ."
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  FOR zz& = 1 TO (LOF(1) / UBCustRecLen)
    GET UBFile, zz&, UBCustRec(1)
    IF UBCustRec(1).Book = "10" THEN
      TCnt = TCnt + 1
      UBCustRec(1).BillCycl = 1
'      FOR zzcnt = 1 TO 7
'        Prev$ = QPTrim$(STR$(UBCustRec(1).LocMeters(zzcnt).PrevRead))
'        Bit$ = LEFT$(Prev$, 2)
'        SELECT CASE Bit$
'        CASE "98", "99"
'          UBCustRec(1).LocMeters(zzcnt).CurRead = UBCustRec(1).LocMeters(zzcnt).PrevRead
'          UBCustRec(1).LocMeters(zzcnt).ReadFlag = "Y"
'          Fixed = Fixed + 1
'        END SELECT
'        UBCustRec(1).LocMeters(zzcnt).ReadFlag = "N"
'      NEXT
      PUT UBFile, zz&, UBCustRec(1)
    END IF
  NEXT

  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing complete."
  PRINT "Changed:"; TCnt, Fixed
  CALL KillFile("fixem.exe")

