DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  CLS
  PRINT "Processing Customers. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen

  NumCustRec& = LOF(CustFile) \ UBCustRecLen
  LOCATE 3, 1: PRINT "Pass 1 of 2"
  FOR LCnt = 1 TO NumCustRec&
    LOCATE 5, 1: PRINT "Processing: "; LCnt; " of "; NumCustRec&
    GET CustFile, LCnt, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      DidThem = 0
      FOR RCnt = 1 TO 15
        IF LEN(QPTrim$(UBCustRec(1).Serv(RCnt).RMTRTYPE)) > 0 THEN
          IF LEN(QPTrim$(UBCustRec(1).Serv(RCnt).RateCode)) = 0 THEN
            UBCustRec(1).Serv(RCnt).RMTRTYPE = ""
            DidThem = -1
            'PRINT LCnt, RCnt
            Bad = Bad + 1
          END IF
        END IF
      NEXT
      IF DidThem THEN
        PUT CustFile, LCnt, UBCustRec(1)
      END IF
    END IF
  NEXT


  LOCATE 3, 1: PRINT "Pass 2 of 2"
  FOR LCnt = 1 TO NumCustRec&
    LOCATE 5, 1: PRINT "Processing: "; LCnt; " of "; NumCustRec&
    GET CustFile, LCnt, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      IF LEN(QPTrim$(UBCustRec(1).Serv(1).RMTRTYPE)) > 0 THEN
        IF LEN(QPTrim$(UBCustRec(1).Serv(2).RMTRTYPE)) = 0 THEN
          UBCustRec(1).Serv(1).RMTRTYPE = "W"
          UBCustRec(1).LocMeters(1).MTRType = "W"
          PUT CustFile, LCnt, UBCustRec(1)
          WBad = WBad + 1
        END IF
      END IF
      IF LEN(QPTrim$(UBCustRec(1).Serv(2).RMTRTYPE)) > 0 THEN
        IF LEN(QPTrim$(UBCustRec(1).Serv(1).RMTRTYPE)) = 0 THEN
          UBCustRec(1).Serv(2).RMTRTYPE = "S"
          UBCustRec(1).LocMeters(1).MTRType = "S"
          PUT CustFile, LCnt, UBCustRec(1)
          SBad = SBad + 1
        END IF
      END IF
    END IF
  NEXT
  
  CLOSE

PRINT
PRINT
PRINT "Corrected:"; WBad; " Water, "; SBad; " Sewer"

