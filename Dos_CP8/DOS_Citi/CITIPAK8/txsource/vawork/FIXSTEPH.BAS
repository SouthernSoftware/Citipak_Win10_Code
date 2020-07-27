DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PROPAbst.BI'

  CONST TaxPersFile = "TAXPERS.DAT"
  Fil$ = SPACE$(7)


  CLS
  PRINT "Processing. . . "
  REDIM PersRec(1)  AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  FOR Cnt& = 1 TO LOF(PersFile) / PersRecLen
    GET #PersFile, Cnt&, PersRec(1)
    IF PersRec(1).EXMPSENI > 0 THEN
      STOP
    END IF
    'PYear$ = RIGHT$(PersRec(1).DESC5, 7)
    'PYear$ = QPTrim$(PYear$)
    'TYear$ = (RIGHT$(PYear$, 4) + " ")
    'RSET Fil$ = TYear$
    'MID$(PersRec(1).DESC5, 25) = Fil$

'    IF VAL(PYear$) = 2001 THEN
'      PersRec(1).LastYrPrinted = 2000
'      dodo = dodo + 1
     ' PUT #PersFile, Cnt&, PersRec(1)
    'END IF
  NEXT
  CLOSE
  CALL KillFile("fixsteph.exe")
  LOCATE 12, 1

  PRINT "Done."
'  PRINT "Changed:"; dodo
  END

