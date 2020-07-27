DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))
  CLS
  LOCATE 2, 1, 0

  PRINT "Processing Billing Cycles. . ."

  OPEN "rate.lst" FOR OUTPUT AS #10
  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    'LOCATE 5, 1
    'PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    'Rate1$ = QPTrim$(UBCustRec(1).Serv(1).RateCode)
    Rate1$ = QPTrim$(UBCustRec(1).Serv(9).RateCode)
    Pos1 = INSTR("LA2LMALMWLMW2", Rate1$)
    'Pos2 = INSTR("LA2LMALMWLMW2", Rate2$)

    IF LEN(Rate1$) > 0 THEN 'OR LEN(Rate2$) > 0 THEN
      IF Pos1 > 0 THEN 'OR Pos2 > 0 THEN
      'PRINT #10, Cnt&, Rate1$, Rate2$
        PRINT Cnt&, Rate1$, Rate2$
        'STOP
      END IF
    END IF


  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Processing Complete."
  PRINT

