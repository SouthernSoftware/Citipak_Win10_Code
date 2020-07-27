DEFINT A-Z
DIM zz AS STRING

OPEN "export.txt" FOR INPUT AS #1
OPEN "export.fix" FOR OUTPUT AS #2

DO WHILE NOT EOF(1)
LINE INPUT #1, zz$
zz$ = zz$ + " "
PRINT #2, zz$
LOOP
CLOSE



