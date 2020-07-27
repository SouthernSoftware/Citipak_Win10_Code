DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Scanning Customers. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 1 TO NumOfRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; " of"; NumOfRecs;
    GET #UBFile, Cnt, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF UBCustRec(1).Status <> "A" THEN
        IF uRound#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) <> 0 THEN
          IF UBCustRec(1).Status <> "B" THEN
            Fixed = Fixed + 1
            UBCustRec(1).Status = "B"
            PUT #UBFile, Cnt, UBCustRec(1)
          END IF
        END IF
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Fixed:"; Fixed

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

