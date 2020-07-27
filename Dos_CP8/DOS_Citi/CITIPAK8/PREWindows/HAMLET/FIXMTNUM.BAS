DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  CONST False = 0, True = NOT False

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  LOCATE 2, 1, 0
  PRINT "Processing Accounts. . ."

  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = False
    FOR Mtr = 1 TO 7
      MtrNum$ = QPTrim$(UBCustRec(1).LocMeters(Mtr).MtrNum)
      IF LEN(MtrNum$) > 0 THEN
        PndPos = INSTR(MtrNum$, "#")
        DO WHILE PndPos > 0
          IF PndPos > 1 THEN
            MtrNum$ = LEFT$(MtrNum$, PndPos - 1) + MID$(MtrNum$, PndPos + 1)
          ELSE
            MtrNum$ = MID$(MtrNum$, PndPos + 1)
          END IF
          PndPos = INSTR(MtrNum$, "#")
        LOOP
        UBCustRec(1).LocMeters(Mtr).MtrNum = MtrNum$
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Complete."
  PRINT
  PRINT

