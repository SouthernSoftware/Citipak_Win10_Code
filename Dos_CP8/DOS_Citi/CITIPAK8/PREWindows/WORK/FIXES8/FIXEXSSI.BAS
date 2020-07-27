DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

TYPE UBSensusGetReadRecType     ' File Layout For Retreiving Records
    Account AS STRING * 8
    NotUsed AS STRING * 5
    MeterID AS STRING * 8
    Filler AS STRING * 1
    Reading AS STRING * 8
    NotUse1 AS STRING * 1
    DateRead AS STRING * 6
    NotUse2 AS STRING * 4
END TYPE

  REDIM UBSenGetRdRec(1) AS UBSensusGetReadRecType
  UBSenGetRdRecLen = LEN(UBSenGetRdRec(1))

  OPEN "c:\stars\exssi002.dat" FOR RANDOM AS #1 LEN = UBSenGetRdRecLen
  FOR Cnt = 1 TO LOF(1) \ UBSenGetRdRecLen
  GET #1, Cnt, UBSenGetRdRec(1)
  WhatDate$ = QPTrim$(UBSenGetRdRec(1).DateRead)
  WhatDate$ = WhatDate$ + "00"
  LSET UBSenGetRdRec(1).DateRead = WhatDate$
  PUT #1, Cnt, UBSenGetRdRec(1)
  NEXT
CLOSE

