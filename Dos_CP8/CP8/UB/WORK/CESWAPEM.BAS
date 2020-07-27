DEFINT A-Z
TYPE CERecType
  ATxt AS STRING * 175
END TYPE

DIM RecLen AS INTEGER
DIM CeRec AS CERecType
DIM RecCnt AS INTEGER
DIM SwapEm(1 TO 2) AS STRING
DIM cnt AS INTEGER
RecLen = LEN(CeRec)
CLS

OPEN "cemtread.dat" FOR RANDOM AS #1 LEN = RecLen
RecCnt = LOF(1) / RecLen

  FOR cnt = 1 TO RecCnt
    GET #1, cnt, CeRec
    SwapEm(1) = MID$(CeRec.ATxt, 72, 7)
    SwapEm(2) = MID$(CeRec.ATxt, 155, 7)
    SWAP SwapEm(1), SwapEm(2)
    MID$(CeRec.ATxt, 72, 7) = SwapEm(1)
    MID$(CeRec.ATxt, 155, 7) = SwapEm(2)
    PUT #1, cnt, CeRec
  NEXT
CLOSE
PRINT
PRINT "CE_Swap completed."

