DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'PROPAbst.BI'

  DIM SHARED PersRec(1 TO 2) AS PersonalRecType

  TaxFile = FREEFILE

  CLS
  PRINT "Processing. . . "

  OPEN "TAXPERS.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(TaxFile) / LEN(PersRec(1))
  FOR Cnt& = 1 TO NumOfPersRecs
    GET TaxFile, Cnt&, PersRec(1)
    IF PersRec(1).MCVALUE > 0 THEN
'      LPRINT PersRec(1).CustPin
      PersRec(1).PERSVAL = PersRec(1).PERSVAL + PersRec(1).MCVALUE
      PersRec(1).MCVALUE = 0
      FdCnt = FdCnt + 1
      PUT TaxFile, Cnt&, PersRec(1)
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  KILL "FIXBROOK.EXE"

