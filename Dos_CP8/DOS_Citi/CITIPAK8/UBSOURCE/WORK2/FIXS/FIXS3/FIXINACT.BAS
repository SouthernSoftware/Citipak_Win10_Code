DECLARE FUNCTION uRound# (N#)
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
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
    IF UBCustRec(1).Status = "B" OR UBCustRec(1).Status = "I" THEN
      'UBCustRec(1).Book = ""
      'UBCustRec(1).SeqNumb = ""
      FOR RDCnt = 1 TO 7
        UBCustRec(1).LocMeters(RDCnt).PrevRead = UBCustRec(1).LocMeters(RDCnt).CurRead
      NEXT
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished."
PRINT "You will need to REINDEX to complete this operation."

