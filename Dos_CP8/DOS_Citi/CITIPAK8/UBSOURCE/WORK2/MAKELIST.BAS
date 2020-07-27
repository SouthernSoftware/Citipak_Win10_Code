DEFINT A-Z
OPEN "ublog.dat" FOR INPUT AS #1
DO UNTIL EOF(1)
  LINE INPUT a$
  if instr$(a$,"Entering

