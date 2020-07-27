DEFINT A-Z


DIM tl$
DIM a$
DIM thisp AS INTEGER

OPEN "prtst_tc.txt" FOR INPUT AS #1
OPEN "timeclk.txt" FOR OUTPUT AS #2

DO
LINE INPUT #1, a$

thisp = INSTR(a$, "-")
tl$ = LEFT$(a$, thisp - 3) + MID$(a$, thisp + 9)
PRINT a$
PRINT #2, tl$

LOOP UNTIL EOF(1)

CLOSE
END

