DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  ReadDate = Date2Num("10-19-2001")

  CLS

  PRINT
  PRINT
  PRINT "Processing Customers. . ."
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
    SELECT CASE UBCustRec(1).Book
    CASE "16", "17", "36", "37"
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 THEN
          CurRead& = UBCustRec(1).LocMeters(MCnt).CurRead
          PreRead& = UBCustRec(1).LocMeters(MCnt).PrevRead
          ConRead& = CurRead& - PreRead&
          UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).CurRead
          UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).CurRead + ConRead&
          UBCustRec(1).LocMeters(MCnt).CurDate = ReadDate
        END IF
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
    END SELECT
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
KillFile "FIXLANDS.EXE"

