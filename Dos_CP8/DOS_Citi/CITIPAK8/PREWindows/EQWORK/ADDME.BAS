DECLARE FUNCTION MakeBucks$ (BucksIn#)

DEFINT A-Z
FUNCTION MakeBucks$ (BucksIn#)

DIM BucksOut$, Cash$, Cent$
DIM cPos AS INTEGER

IF BucksIn# = 0 THEN
  BucksOut$ = "0.00"
  GOTO BucksExit
END IF

BucksOut$ = QPTrim(STR$(Round(BucksIn#)))

cPos = INSTR(BucksOut$, ".")

IF cPos > 1 THEN   'format is ?.? possiable ????.??
  Cash$ = LEFT$(BucksOut$, cPos - 1)
  Cent$ = MID$(BucksOut$, cPos)
ELSEIF cPos = 1 THEN ' format is .?"
  Cash$ = "0"
  Cent$ = BucksOut$
ELSE 'no cents
  Cash$ = BucksOut$
  Cent$ = ".00"
END IF

IF LEN(Cent$) = 2 THEN 'format is .5 must append 0
  Cent$ = Cent$ + "0"
END IF

BucksOut$ = Cash$ + Cent$

BucksExit:

  MakeBucks$ = BucksOut$

END FUNCTION

