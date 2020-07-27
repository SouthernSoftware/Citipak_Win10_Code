DEFINT A-Z

'$INCLUDE: 'logicon.txt'


DIM elr AS UBLogiconReadRecType
DIM Ilr AS UBLogiconGetReadRecType
DIM RecCnt AS LONG
DIM CurRead AS LONG
DIM PreRead AS LONG
DIM AvgRead AS LONG
DIM cnt AS LONG
DIM twrk AS STRING * 8
DIM MKRead AS STRING
DIM VooDoo1 AS STRING
DIM VooDoo2 AS STRING

CLS

twrk = "00000000"

VooDoo1 = "091030"
VooDoo2 = "104308"
OPEN "C:\wblogno5.dat" FOR RANDOM AS #1 LEN = 256
RecCnt = LOF(1) / 256

  FOR cnt = 1 TO RecCnt
    GET #1, cnt, elr
    IF elr.rectype = "A" THEN
      PreRead = VAL(elr.PrevRead)
      AvgRead = VAL(elr.AvgUse)
      MKRead = RIGHT$(twrk + LTRIM$(RTRIM$(STR$(PreRead + AvgRead))), 8)
      elr.ReadDate = VooDoo1
      elr.ReadTime = VooDoo2
      elr.CurRead = MKRead
      PUT #1, cnt, elr

    END IF
  NEXT
CLOSE

