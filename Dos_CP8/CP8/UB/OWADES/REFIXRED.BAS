DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (Text$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'
  
  CONST True = -1, False = 0

  T1Date = Date2Num("04-30-1997")
  T2Date = Date2Num("05-30-1997")
  T3Date = Date2Num("06-30-1997")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  PRINT
  PRINT

  PRINT
  PRINT
  PRINT "Adjusting Readings"
  PRINT
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  NumOfRecs& = LOF(UBTran) \ UBTranRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBTran, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE T3Date
        GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        IF INSTR(UBCustRec(1).CUSTTYPE, "MON") <= 0 THEN
          SELECT CASE UBCustRec(1).BOOK
          CASE "03", "04"  '063097
            FOR MtrCnt = 1 TO 7
              UBCustRec(1).LocMeters(MtrCnt).CurRead = UBTranRec(1).CurRead(MtrCnt)
              UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBTranRec(1).PrevRead(MtrCnt)
            NEXT
            PUT UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
          END SELECT
        END IF
      CASE T1Date
        GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        IF INSTR(UBCustRec(1).CUSTTYPE, "MON") <= 0 THEN
          SELECT CASE UBCustRec(1).BOOK
          CASE "01", "07", "08", "09", "10"    '043097
            FOR MtrCnt = 1 TO 7
              UBCustRec(1).LocMeters(MtrCnt).CurRead = UBTranRec(1).CurRead(MtrCnt)
              UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBTranRec(1).PrevRead(MtrCnt)
            NEXT
            PUT UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
          END SELECT
        END IF
      CASE T2Date
        GET UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        IF INSTR(UBCustRec(1).CUSTTYPE, "MON") <= 0 THEN
          SELECT CASE UBCustRec(1).BOOK
          CASE "02", "05", "06", "11" ' 053097
            FOR MtrCnt = 1 TO 7
              UBCustRec(1).LocMeters(MtrCnt).CurRead = UBTranRec(1).CurRead(MtrCnt)
              UBCustRec(1).LocMeters(MtrCnt).PrevRead = UBTranRec(1).PrevRead(MtrCnt)
            NEXT
            PUT UBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
          END SELECT
        END IF
      END SELECT
    END IF

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished!"

