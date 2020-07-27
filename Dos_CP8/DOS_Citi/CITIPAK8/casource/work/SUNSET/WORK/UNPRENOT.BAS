DECLARE SUB ShowProcessing (Cnt&, NumOfRecs&)
DEFINT A-Z

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  PRINT "   Clear prenote Flags"
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    ShowProcessing Cnt&, CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).PreNoteFlag <> 0 THEN
      UBCustRec(1).PreNoteFlag = 0
      TCnt = TCnt + 1
      'PUT UBCFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "      Changed: "; TCnt

