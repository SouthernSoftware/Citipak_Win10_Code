DEFINT A-Z

  '$INCLUDE: 'PROPAbst.BI'
  ''$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'

  REDIM PersRec(1) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))
  
  CLS
  PRINT "         WARNING WARNING WARNING WARNING WARNING WARNING WARNING!!!"
  PRINT
  PRINT "   ALL OTHER TAX OPERATION MUST HALT UNTIL THIS PROCEDURE HAS COMPLETED!!!"
  PRINT
  PRINT
  PRINT
  PRINT "   CLEAR ALL PERSONAL PROPERTY VALUATIONS."
  PRINT
  LINE INPUT "   ARE YOU SURE (YES/NO)? "; YesNo$
  PRINT
  PRINT
  IF UCASE$(YesNo$) <> "YES" THEN
    PRINT "   CLEAR VALUATIONS ABORTED."
    PRINT
    PRINT
    GOTO DontDoItExit
  END IF
  
  PRINT "   Clearing Personal Property Valuations"
  PRINT

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  NumRecs& = LOF(PersFile) / PersRecLen
  FOR Cnt& = 1 TO NumRecs&
    LOCATE , 1
    PRINT "   Processing:"; Cnt&; "of"; NumRecs&;
    GET #PersFile, Cnt&, PersRec(1)
    PersRec(1).PersVal = 0
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    PersRec(1).LATELIST = "N"
    PUT #PersFile, Cnt&, PersRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "   Processing Complete."

DontDoItExit:

  END

