DEFINT A-Z
CLS
OPEN "value.txt" FOR INPUT AS #1
DO WHILE NOT EOF(1)
  LINE INPUT #1, a$
  IF INSTR(a$, "CEMETERY") > 0 THEN
    FoundCnt = FoundCnt + 1
  END IF
LOOP
CLOSE
PRINT "Found:"; FoundCnt


