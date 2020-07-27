DEFINT A-Z

'This

  TabChar$ = CHR$(9)
  Space8$ = ","  'SPACE$(8)

  OPEN "txhst01.txt" FOR INPUT AS #1 LEN = 512
  OPEN "Ztxhst01.txt" FOR OUTPUT AS #2 LEN = 512

  DO
    LINE INPUT #1, Text$
    WhatLen = LEN(Text$)
    TabPos = INSTR(Text$, TabChar$)
    DO WHILE TabPos > 0
      Text$ = LEFT$(Text$, TabPos - 1) + Space8$ + MID$(Text$, TabPos + 1)
      TabPos = INSTR(Text$, TabChar$)
    LOOP
    PRINT #2, Text$
  LOOP UNTIL EOF(1)
  CLOSE

