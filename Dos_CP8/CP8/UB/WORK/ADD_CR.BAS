DEFINT A-Z


DIM Test AS STRING * 194
DIM Tst2 AS STRING * 192

DIM Mess$
DIM RecCnt&, lopCnt&, x&

GOTO DownHere


OPEN "ssi001.rte" FOR RANDOM AS #1 LEN = 192
OPEN "ssi001CR.rte" FOR OUTPUT AS #2
RecCnt& = LOF(1) \ 192
FOR x& = 1 TO RecCnt&
 GET #1, , Test$
 PRINT #2, Test$
NEXT

'PRINT RecCnt
CLOSE

DownHere:
KILL "EXSSI001.000"

OPEN "EXSSI001.rte" FOR INPUT AS #1

OPEN "EXSSI001.000" FOR RANDOM AS #2 LEN = 192

RecCnt& = LOF(1) \ 194
'PRINT RecCnt&
'END
FOR x& = 1 TO RecCnt&
 LINE INPUT #1, Mess$
 'GET #1, x&, Test$
 'PRINT LEFT$(Mess$, 20)
 Tst2$ = Mess$
 'PRINT Tst2
 PUT #2, , Tst2$
 'SLEEP
NEXT
 CLOSE
 END

'

