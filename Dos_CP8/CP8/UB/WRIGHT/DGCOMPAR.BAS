DEFINT A-Z
DECLARE FUNCTION QPTrim$ (thetext$)

OPEN "dodo.txt" FOR INPUT AS #1
DO UNTIL EOF(1)
  LINE INPUT #1, a$
  b$ = LEFT$(a$, 28)
  c$ = MID$(a$, 29)
  b$ = QPTrim$(b$)
  c$ = QPTrim$(c$)
  IF b$ <> c$ THEN
    LPRINT b$; TAB(40); c$
  END IF
LOOP
CLOSE

