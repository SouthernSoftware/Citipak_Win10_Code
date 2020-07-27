DEFINT A-Z
OPEN "fdtax1.frm" FOR BINARY AS #1
a$ = SPACE$(LOF(1))
GET #1, , a$
CLOSE

OPEN "fdtax2.frm" FOR BINARY AS #1
'a$ = SPACE$(LOF(1))
'GET #1, , a$
DO
 RepPos = INSTR(a$, "FTS")
 IF RepPos > 0 THEN
   MID$(a$, RepPos, 3) = "FTM"
 END IF
 LOOP WHILE RepPos > 0

PUT #1, , a$
CLOSE

