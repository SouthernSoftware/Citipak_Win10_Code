DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  PRINT
  PRINT
  PRINT "Processing AvgUse. . ."
  PRINT
  PRINT

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)

    REDIM TotalUse(1 TO 7) AS LONG
    REDIM UseCnt(1 TO 7) AS INTEGER

    LastTran& = UBCustRec(1).LastTrans
    DO WHILE LastTran& > 0
      GET UBTran, LastTran&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill THEN
        FOR MCnt = 1 TO 7
          IF UBTranRec(1).CurRead(MCnt) > 0 THEN
            TotalUse(MCnt) = TotalUse(MCnt) + (UBTranRec(1).CurRead(MCnt) - UBTranRec(1).PrevRead(MCnt))
            UseCnt(MCnt) = UseCnt(MCnt) + 1
          END IF
        NEXT
      END IF
      LastTran& = UBTranRec(1).PrevTrans
    LOOP
    FOR MCnt = 1 TO 7
      IF TotalUse(MCnt) > 0 THEN
        UBCustRec(1).LocMeters(MCnt).AvgUse = TotalUse(MCnt) / UseCnt(MCnt)
        UBCustRec(1).LocMeters(MCnt).UseCnt = UseCnt(MCnt)
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done."

