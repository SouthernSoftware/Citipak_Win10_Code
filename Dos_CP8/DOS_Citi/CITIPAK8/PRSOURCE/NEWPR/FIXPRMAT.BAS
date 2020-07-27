DEFINT A-Z
'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRTRANS.Bi'
'$INCLUDE: 'oPRtypes.Bi'
'$INCLUDE: 'PRFiles.bi'

DIM Trans  AS TransRecType
TranLen = LEN(Trans)

OPEN TransHistFileName FOR RANDOM AS #1 LEN = TranLen
GET #1, 2870, Trans
Trans.MatchRetAmt = 0
PUT #1, 2870, Trans
CLOSE

KILL "FIXPRMAT.EXE"

