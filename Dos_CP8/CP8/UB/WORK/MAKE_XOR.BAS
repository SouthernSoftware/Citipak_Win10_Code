DEFINT A-Z
a$ = "TOWN OF HOLLY SPRINGS"
CLS

OPEN "holly.txt" FOR OUTPUT AS #2
FOR zz = 1 TO LEN(a$)
  ch = ASC(MID$(a$, zz, 1))
  ch = (ch XOR 126) XOR zz
  PRINT #2, CHR$(ch);
  PRINT CHR$((ch XOR 126) XOR zz);
NEXT

CLOSE

