DEFINT A-Z
TYPE CEFillType
  CELine AS STRING * 175
END TYPE

DIM CERec AS CEFillType

RecLen = LEN(CERec)

OPEN "CEMTREAD.DAT" FOR RANDOM AS #1 LEN = RecLen
NumOfRec = LOF(1) / RecLen
FOR cnt = 1 TO NumOfRec 'STEP 300
  GET #1, cnt, CERec
  MID$(CERec.CELine, 79, 1) = "Y"
  PUT #1, cnt, CERec
NEXT
  GET #1, 123, CERec
  MID$(CERec.CELine, 79, 1) = "N"
  PUT #1, 123, CERec

  GET #1, 775, CERec
  MID$(CERec.CELine, 79, 1) = "N"
  PUT #1, 775, CERec

  GET #1, 1531, CERec
  MID$(CERec.CELine, 79, 1) = "N"
  PUT #1, 1531, CERec
CLOSE

