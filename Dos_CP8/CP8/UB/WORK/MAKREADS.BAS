DEFINT A-Z
TYPE dodo
  a AS STRING * 175
END TYPE
DIM arec AS dodo

DIM nc AS STRING * 9

RecLen = LEN(arec)

OPEN "cemtread.dat" FOR RANDOM AS #1 LEN = RecLen

RecCnt = LOF(1) / RecLen

FOR Cnt = 1 TO RecCnt
  GET #1, Cnt, arec
  Prev$ = LTRIM$(RTRIM$(MID$(arec.a, 62, 9)))
  a& = VAL(Prev$)
  b& = a& + 1500
  LSET nc = STR$(b&)
  PRINT a&, , nc
  MID$(arec.a, 92, 9) = nc
  MID$(arec.a, 79, 1) = "Y"
  MID$(arec.a, 101, 14) = "09222009114723"
  PUT #1, Cnt, arec

NEXT
CLOSE

