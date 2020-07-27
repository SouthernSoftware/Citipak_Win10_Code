DEFINT A-Z
OPEN "UBPREFIN.RPT" FOR INPUT AS #1
OPEN "DEPOSIT.TXT" FOR OUTPUT AS #2
DO UNTIL EOF(1)
  LINE INPUT #1, a$
  IF LEFT$(a$, 1) = "F" THEN
    b$ = a$
    'PRINT #2, a$
  END IF
  IF INSTR(a$, "Deposit") > 0 THEN
    none$ = RIGHT$(a$, 5)
    IF none$ <> "$0.00" THEN
      PRINT #2, b$
      PRINT #2, a$
    END IF
  END IF
LOOP
CLOSE

