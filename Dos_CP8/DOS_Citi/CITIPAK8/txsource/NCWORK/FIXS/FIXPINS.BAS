DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  PRINT "Processing. . ."
  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

PropRestart:

  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  NumPropRecs& = LOF(PropFile) / PropRecLen

  FOR WhatProp& = 1 TO NumPropRecs&
    GET PropFile, WhatProp&, PropRec(1)
    ThisPin$ = QPTrim$(PropRec(1).REALPIN)
    IF LEN(ThisPin$) > 0 THEN
      PinChar$ = LEFT$(ThisPin$, 1)
      SELECT CASE PinChar$
      CASE "7", "9"
        ThisPin$ = "0" + ThisPin$
        PropRec(1).REALPIN = ThisPin$
        Fixed = Fixed + 1
        PUT PropFile, WhatProp&, PropRec(1)
      END SELECT
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL killfile("fixpins.exe")
  END

