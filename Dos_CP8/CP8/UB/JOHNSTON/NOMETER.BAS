DEFINT A-Z

DECLARE FUNCTION QPTrim$ (Text$)


  '$INCLUDE: 'Newcust.bi'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfRecs& = LOF(UBCust) \ UBCustRecLen

  FOR CustAcctNo& = 1 TO NumOfRecs&
    GET UBCust, CustAcctNo&, UBCustRec(1)
    IF UBCustRec(1).Status = "A" THEN
      FOR mChk = 1 TO 7
        MTRtype$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRtype)
        IF LEN(MTRtype$) > 0 THEN

          MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRNUM)
          IF LEN(MtrNumb$) = 0 THEN
            Found = Found + 1
            'print this cust rec here
            EXIT FOR
          END IF
        END IF
      NEXT
    END IF
  NEXT
  CLOSE
  PRINT "Found:"; Found

