DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  CLS

  REDIM TaxCust(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCust(1))

  OPEN "TaxCust.DAT" FOR RANDOM SHARED AS #1 LEN = TaxCustRecLen

  NumOfRec& = LOF(1) / TaxCustRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxCust(1)
      CName$ = QPTrim$(TaxCust(1).LName)
      LLen = LEN(CName$)
      IF LLen > 2 THEN
        FOR zz = LLen TO 1 STEP -1
          ThisChar$ = MID$(CName$, zz, 1)
          IF ThisChar$ = " " THEN
            TaxCust(1).FName = MID$(CName$, 1, zz - 1)
            TaxCust(1).LName = MID$(CName$, zz + 1)

            SSName$ = QPTrim$(TaxCust(1).LName)
            FFName$ = QPTrim$(TaxCust(1).FName)
            SFName$ = SSName$ + FFName$
            SpacePos = INSTR(SFName$, " ")
            DO WHILE SpacePos > 0
              SFName$ = LEFT$(SFName$, SpacePos - 1) + MID$(SFName$, SpacePos + 1)
              SpacePos = INSTR(SFName$, " ")
            LOOP

            TaxCust(1).SName = SFName$
            EXIT FOR
          END IF
        NEXT
      END IF
    PUT #1, Cnt&, TaxCust(1)
  NEXT
  CLOSE
  PRINT
  PRINT

  PRINT "Done."
  KILL "FIXCUST.EXE"

