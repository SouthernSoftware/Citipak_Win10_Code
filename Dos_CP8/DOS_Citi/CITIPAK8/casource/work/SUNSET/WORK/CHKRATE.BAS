DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing CutOff Flags. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    GET UBCFile, Cnt&, UBCustRec(1)
    TotBal# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
    IF TotBal# < 0 THEN
      UBCustRec(1).Status = "I"
      PUT UBCFile, Cnt&, UBCustRec(1)
    END IF
'    IF UBCustRec(1).Monthly(1).PayAmt > 0 THEN
'      STOP
'    END IF
'    IF UBCustRec(1).Monthly(2).PayAmt > 0 THEN
'      STOP
'    END IF

'    Rat1$ = QPTrim$(UBCustRec(1).Serv(1).RateCode)
'    Rat2$ = QPTrim$(UBCustRec(1).Serv(2).RateCode)
'    Rat3$ = QPTrim$(UBCustRec(1).Serv(3).RateCode)
'    IF LEN(Rat1$) > 0 THEN
'      IF LEFT$(Rat1$, 1) <> "W" THEN
'        STOP
'      END IF
'    END IF
'    IF LEN(Rat2$) > 0 THEN
'      IF LEFT$(Rat2$, 1) <> "S" THEN
'        STOP
'      END IF
'    END IF
'    IF LEN(Rat3$) > 0 THEN
'      IF LEFT$(Rat3$, 1) <> "G" THEN
'        STOP
'      END IF
'    END IF



'    UBCustRec(1).CUTOFFYN = "Y"
    'PUT UBCFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIXCUTOF.EXE"

