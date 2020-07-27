DEFINT A-Z

DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE SUB KillFile (FileName$)


  '$INCLUDE: 'GL.bi'                        'GL File Types
  DIM SHARED Trans AS GLTransRecType

  BDate1 = Date2Num("01-01-1900")
  BDate2 = Date2Num("12-31-1900")

  CLS
  PRINT "Correcting Transaction Dates. . ."
  TransRecLen = LEN(Trans)
  TransFileNum = FREEFILE
  OPEN "GLTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TransFileNum LEN = TransRecLen
  NumTrans& = LOF(TransFileNum) \ TransRecLen
  FOR Cnt& = 1 TO NumTrans&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumTrans&;
    GET TransFileNum, Cnt&, Trans
    IF Trans.TRDate >= BDate1 AND Trans.TRDate <= BDate2 THEN
      WhatDate$ = Num2Date(Trans.TRDate)
      WhatDate$ = LEFT$(WhatDate$, 6) + "2000"
      Trans.TRDate = Date2Num(WhatDate$)
      PUT TransFileNum, Cnt&, Trans
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Corrected:"; Fixed
  KillFile "fixgldat.exe"

