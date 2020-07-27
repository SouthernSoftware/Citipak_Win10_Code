DEFINT A-Z
TYPE CeRecType
  SomeStuff  AS STRING * 100
  TDate      AS STRING * 14
  MoreStuff  AS STRING * 61
END TYPE

DIM CERec AS CeRecType
PRINT LEN(CERec)
OPEN "cemtread.dat" FOR RANDOM AS #1 LEN = 175
NumRec = LOF(1) / 175
FOR zz = 1 TO NumRec
  GET #1, zz, CERec
  CERec.TDate = "10142009120001"
  PUT #1, zz, CERec
NEXT
CLOSE

