TYPE ImpExpReadType
  CurrRead AS LONG
  PrevRead AS LONG
  ReadDate AS INTEGER
  ReadFlag AS STRING * 1
END TYPE

TYPE SewImpExpRecType
  Location AS STRING * 8
  Reads(1 TO 7) AS ImpExpReadType
END TYPE

