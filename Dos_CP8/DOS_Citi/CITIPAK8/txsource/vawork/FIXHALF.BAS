DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PROPAbst.BI'

  CONST TaxPersFile = "TAXPERS.DAT"

  CLS
  PRINT "Processing. . . "
  REDIM PersRec(1)  AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  FOR Cnt& = 1 TO LOF(PersFile) / PersRecLen
    GET #PersFile, Cnt&, PersRec(1)
    PYear$ = RIGHT$(PersRec(1).DESC5, 5)
    PYear$ = QPTrim$(PYear$)
    IF VAL(PYear$) >= 2000 THEN
      MID$(PersRec(1).DESC5, 27) = "2002"
      PersRec(1).LastYrPrinted = 2001
      dodo = dodo + 1
      PUT #PersFile, Cnt&, PersRec(1)
    END IF
  NEXT
  CLOSE
  CALL KillFile("fixhalf.exe")
  LOCATE 12, 1

  PRINT "Done."
  PRINT "Changed:"; dodo
  END

