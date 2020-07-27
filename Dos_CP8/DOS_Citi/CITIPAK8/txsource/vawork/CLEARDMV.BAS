DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PROPAbst.BI'

  CONST TaxPersFile = "TAXPERS.DAT"

  WhatYear$ = COMMAND$

  IF LEN(WhatYear$) = 0 THEN END
  ThisYear = VAL(WhatYear$)

  CLS
  PRINT "Processing. . . "
  REDIM PersRec(1)  AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  PersRecCnt& = LOF(PersFile) / PersRecLen

  FOR Cnt& = 1 TO PersRecCnt&
    GET #PersFile, Cnt&, PersRec(1)
    PersRec(1).DMVSubmitted = ""
    PYear$ = RIGHT$(PersRec(1).DESC5, 5)
    PYear$ = QPTrim$(PYear$)
    IF VAL(PYear$) = ThisYear THEN
      PersRec(1).DMVSubmitted = ""
      dodo = dodo + 1
      PUT #PersFile, Cnt&, PersRec(1)
    END IF
  NEXT
  CLOSE
  'CALL KillFile("cleardmv.exe")

  LOCATE 12, 1

  PRINT "Done."
  PRINT "Changed:"; dodo
  END

