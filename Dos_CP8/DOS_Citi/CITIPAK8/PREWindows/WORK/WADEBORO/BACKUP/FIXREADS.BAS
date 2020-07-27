DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  TDate = Date2Num("04-01-1997")

  CLS
  PRINT
  PRINT
  PRINT "Processing Readings. . ."

  REDIM UBCustRec(1)   AS NewUBCustRecType
  REDIM UBTranRec(1)  AS UBTransRecType
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))

  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).STATUS = "A" THEN
      DidEmOne = 0
      SELECT CASE UBCustRec(1).BOOK
        CASE "03", "04"
          LastTran& = UBCustRec(1).LastTrans
          DO WHILE LastTran& > 0
            GET #UBTran, LastTran&, UBTranRec(1)
            IF (UBTranRec(1).TransDate = TDate) AND (UBTranRec(1).TransType = TranUtilityBill) THEN
              DidEmOne = -1
              FOR MCnt = 1 TO 7
                IF UBCustRec(1).LocMeters(MCnt).CurRead >= 0 THEN
                  UBCustRec(1).LocMeters(MCnt).PrevRead = UBTranRec(1).CurRead(MCnt) + UBCustRec(1).LocMeters(MCnt).AvgUse
                  UBCustRec(1).LocMeters(MCnt).CurRead = UBCustRec(1).LocMeters(MCnt).PrevRead + UBCustRec(1).LocMeters(MCnt).AvgUse
                END IF
              NEXT
              EXIT DO
            END IF
            LastTran& = UBTranRec(1).PrevTrans
          LOOP
          IF DidEmOne THEN
            PUT UBFile, Cnt&, UBCustRec(1)
          ELSE
            NGCnt = NGCnt + 1'STOP
          END IF

        CASE "02", "05", "06", "11"
          FOR MCnt = 1 TO 7
            IF UBCustRec(1).LocMeters(MCnt).PrevRead >= 0 THEN
              UBCustRec(1).LocMeters(MCnt).PrevRead = UBCustRec(1).LocMeters(MCnt).PrevRead + UBCustRec(1).LocMeters(MCnt).AvgUse
            END IF
          NEXT
          PUT UBFile, Cnt&, UBCustRec(1)
      END SELECT
    END IF
  NEXT
  
  CLOSE

PRINT
PRINT
'PRINT "NG: "; NGCnt
PRINT "Done."

