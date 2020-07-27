DEFINT A-Z
DECLARE FUNCTION QPTrim$ (text$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.bi'
  CLS
  PRINT
  PRINT
  PRINT "Processing. . ."
  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfRecs = LOF(UBCust) / UBCustRecLen

  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen

  FOR Cnt = 1 TO NumOfRecs
    LOCATE 7, 1
    PRINT "Customer:"; Cnt; "of"; NumOfRecs;
    GET UBCust, Cnt, UBCustRec(1)

    Trans& = UBCustRec(1).LastTrans
    DO WHILE Trans& <> 0
      GET UBTrans, Trans&, UBTrans(1)
      IF UBTrans(1).TransType = TranUtilityBill THEN
        Amount# = UBTrans(1).TransAmt
        TAmount# = 0
        FalCnt = 0
        IF Amount# >= 15 THEN
          TAmount# = Amount# - 15
          FalCnt = 4
          IF TAmount# > 0 THEN
            FalCnt = 4 + (TAmount# / 20)
          END IF
          UBCustRec(1).PUMPCODE = QPTrim$(STR$(FalCnt))
          PUT UBCust, Cnt, UBCustRec(1)
        END IF
      END IF
      Trans& = UBTrans(1).PrevTrans
    LOOP
 NEXT

 CLOSE
 CALL KillFile("setfcnt.exe")

 LOCATE 12, 1
 PRINT "Processing Complete."

