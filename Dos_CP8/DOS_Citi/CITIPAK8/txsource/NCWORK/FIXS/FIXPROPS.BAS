DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  REDIM PropRec(1) AS PropertyRecType
  PropRecLen = LEN(PropRec(1))

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  CLS
  LOCATE 4, 1
  PRINT "Correcting Property Links. . ."
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropRecLen
  NumOfPRop& = LOF(PropFile) / PropRecLen
  FOR Cnt& = 1 TO NumOfPRop&
    LOCATE 7, 1: PRINT "Processing:"; Cnt&; "of"; NumOfPRop&;
    GET PropFile, Cnt&, PropRec(1)
    IF PropRec(1).NextRec > NumOfPRop& THEN
      PropRec(1).NextRec = 0
      PUT PropFile, Cnt&, PropRec(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; Fixed
  CALL KillFile("fixprops.exe")

