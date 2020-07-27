DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.bi'
  CONST False = 0, True = NOT False

  REDIM UBRate(1 TO 1) AS UBRateTblRecType
  RateLen = LEN(UBRate(1))

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  LOCATE 2, 1, 0
  PRINT "Removing Invalid Rate Codes. . ."

  RptFile = FREEFILE
  OPEN "UBBADRAT.DAT" FOR OUTPUT AS RptFile LEN = 255

  UBFile = FREEFILE
  OPEN "ubrate.dat" FOR RANDOM SHARED AS UBFile LEN = RateLen
  NumOfRate& = LOF(UBFile) \ RateLen

  DIM Rates$(1 TO NumOfRate&)

  FOR Cnt& = 1 TO NumOfRate&
    GET UBFile, Cnt&, UBRate(1)
    Rates$(Cnt&) = QPTrim$(UBRate(1).RateCode)
  NEXT
  CLOSE UBFile

  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = False
    TabStop = 1
    FOR Rev = 1 TO 15
      ThisRate$ = QPTrim$(UBCustRec(1).Serv(Rev).RateCode)
      IF LEN(ThisRate$) > 0 THEN
        FoundRate = False
        FOR RCnt = 1 TO NumOfRate&
          IF ThisRate$ = Rates$(RCnt) THEN
            FoundRate = True
            EXIT FOR
          END IF
        NEXT
        IF FoundRate = False THEN
          IF DidEm = False THEN
            PRINT #RptFile, Cnt&;
          END IF
          PRINT #RptFile, TAB(TabStop * 10); Rev;
          TabStop = TabStop + 1
          PRINT #RptFile, TAB(TabStop * 10); UBCustRec(1).Serv(Rev).RateCode;
          UBCustRec(1).Serv(Rev).RateCode = ""
          UBCustRec(1).Serv(Rev).RMtrType = ""
          TabStop = TabStop + 1
          Removed& = Removed& + 1
          DidEm = True
        END IF
      END IF
    NEXT
    IF DidEm THEN
      PUT UBFile, Cnt&, UBCustRec(1)
      DidCnt& = DidCnt& + 1
      PRINT #RptFile,
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Processing Complete."
  PRINT
  PRINT
  PRINT
  PRINT " Accounts Corrected:"; DidCnt&
  PRINT "  Ratecodes Removed:"; Removed&

