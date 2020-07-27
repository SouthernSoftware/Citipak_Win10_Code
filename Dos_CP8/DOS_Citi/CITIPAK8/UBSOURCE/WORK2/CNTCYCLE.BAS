DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION fRound# (N#)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  TDate = Date2Num(DATE$)

  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM Cycles(0 TO 100) AS LONG

  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT "Checking Billing Cycles. . ."
  'PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  'Cycle = 8
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      Active = Active + 1
      IF (UBCustRec(1).BillCycl >= 0) AND (UBCustRec(1).BillCycl <= 100) THEN
        Cycles(UBCustRec(1).BillCycl) = Cycles(UBCustRec(1).BillCycl) + 1
      ELSE
        Bad = Bad + 1
         LPRINT "Invalid Account:"; Cnt&; "  Cycle: ";
        IF UBCustRec(1).BillCycl = -32767 THEN
          LPRINT " NONE!"
        ELSE
          LPRINT UBCustRec(1).BillCycl
        END IF
      END IF
    END IF
  NEXT
  CLOSE UBFile

PRINT
PRINT
FOR Cnt = 0 TO 100
  IF Cycles(Cnt) > 0 THEN
    PRINT "Cycle:"; Cnt; "  Count:"; Cycles(Cnt)
  END IF
NEXT
PRINT
PRINT "  Total Active Accounts:"; Active
PRINT "Active Invalid Accounts:"; Bad
END

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION fRound# (N#)
  fRound# = INT(N# * 100 + .5) / 100
END FUNCTION

