DEFINT A-Z
  DECLARE FUNCTION QPTrim$ (Text$)
  CONST True = -1, False = NOT True

  '$INCLUDE: 'bobTaxes.BI'

  DIM SHARED TaxCustRec(1) AS TaxCustType

  TaxRecLen = LEN(TaxCustRec(1))

  CLS
  OPEN "taxcust.dat" FOR RANDOM AS #1 LEN = TaxRecLen
  RecCnt& = LOF(1) / TaxRecLen
  FOR Cnt& = 1 TO RecCnt&
    GET #1, Cnt&, TaxCustRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; RecCnt&;

    IF LEN(QPTrim$(TaxCustRec(1).SName)) <= 0 THEN
      NGCnt = NGCnt + 1
      GOSUB MakeNewSearchName

      PUT #1, Cnt&, TaxCustRec(1)
    END IF

  NEXT
  LOCATE 8, 1

  PRINT "Short Names Added:"; NGCnt
  CLOSE
  CALL KillFile("fixsname.exe")
  END

MakeNewSearchName:
  TName$ = QPTrim$(TaxCustRec(1).CustName)
  CLen = LEN(TName$)
  SPos = 0
  DO
    SPos = INSTR(TName$, " ")
    IF SPos > 0 THEN
      TName$ = LEFT$(TName$, SPos - 1) + MID$(TName$, SPos + 1)
    ELSE
      SPos = False
    END IF
  LOOP WHILE SPos > False

  SPos = 0
  DO
    SPos = INSTR(TName$, "&")
    IF SPos > 0 THEN
      TName$ = LEFT$(TName$, SPos - 1) + MID$(TName$, SPos + 1)
    ELSE
      SPos = False
    END IF
  LOOP WHILE SPos > False

  SPos = 0
  DO
    SPos = INSTR(TName$, ".")
    IF SPos > 0 THEN
      TName$ = LEFT$(TName$, SPos - 1) + MID$(TName$, SPos + 1)
    ELSE
      SPos = False
    END IF
  LOOP WHILE SPos > False

  SPos = 0
  DO
    SPos = INSTR(TName$, ")")
    IF SPos > 0 THEN
      TName$ = LEFT$(TName$, SPos - 1) + MID$(TName$, SPos + 1)
    ELSE
      SPos = False
    END IF
  LOOP WHILE SPos > False

  SPos = 0
  DO
    SPos = INSTR(TName$, "(")
    IF SPos > 0 THEN
      TName$ = LEFT$(TName$, SPos - 1) + MID$(TName$, SPos + 1)
    ELSE
      SPos = False
    END IF
  LOOP WHILE SPos > False


  TaxCustRec(1).SName = TName$
RETURN

