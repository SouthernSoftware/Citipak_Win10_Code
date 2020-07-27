DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)


  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  CLS
  PRINT
  PRINT "   Coping Pins. ."
  REDIM PropEDRec(1) AS PropertyRecType
  PropEDRecLen = LEN(PropEDRec(1))
  
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropEDRecLen
  NumProp& = LOF(PropFile) \ PropEDRecLen
  FOR Cnt& = 1 TO NumProp&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumProp&;
    GET PropFile, Cnt&, PropEDRec(1)
    TPin$ = UCASE$(PropEDRec(1).PROPNOT3)
    IF INSTR(TPin$, "COUNTY") > 0 THEN
      Pin$ = QPTrim$(MID$(TPin$, 12))
    ELSE
      GOTO NoPinSkip
    END IF
    IF LEN(Pin$) > 0 THEN
      LSET PropEDRec(1).RealPin = Pin$
      Fixed = Fixed + 1
      PUT PropFile, Cnt&, PropEDRec(1)
    END IF
NoPinSkip:
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "   Changed:"; Fixed; "of"; NumProp&
  LOCATE 12, 1
  PRINT "    Done."
  CALL KillFile("movepin.exe")

