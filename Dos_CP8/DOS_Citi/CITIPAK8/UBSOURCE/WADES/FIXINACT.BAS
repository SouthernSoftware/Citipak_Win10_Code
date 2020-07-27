DEFINT A-Z
DECLARE FUNCTION uRound# (N#)
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
      IF uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 THEN
        IF UBCustRec(1).Status = "B" THEN
          UBCustRec(1).Status = "I"
        END IF
      END IF
      Fixed = Fixed + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished:"; Fixed

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

