DEFINT A-Z

TYPE aRecType
  AA AS STRING * 175
END TYPE
DIM RecCnt AS INTEGER
DIM Cnt AS INTEGER
DIM PRead AS STRING
DIM CRead AS STRING
DIM Acct AS STRING

DIM CeRec    AS aRecType
CRecLen = LEN(CeRec)

OPEN "cemtread.dat" FOR RANDOM AS #1 LEN = CRecLen
RecCnt = LOF(1) / CRecLen
FOR Cnt = 1 TO RecCnt
  GET #1, Cnt, CeRec
  CRead = MID$(CeRec.AA, 83, 8)
  PRead = MID$(CeRec.AA, 92, 8)
  Acct = MID$(CeRec.AA, 72, 7)
  MID$(CeRec.AA, 92, 8) = CRead
  MID$(CeRec.AA, 83, 8) = PRead
  MID$(CeRec.AA, 79, 1) = "Y"
  'MID$(CeRec.AA, 155, 7) = Acct
  PUT #1, Cnt, CeRec
NEXT

