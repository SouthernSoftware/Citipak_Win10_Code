DEFINT A-Z

  
  '$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  CLS
  PRINT "Processing. . ."

  REDIM ARCust(1) AS ARCustRecType
'  REDIM ARTransRec(1) AS ARTransRecType

  CALL KillFile("ARTrans.OLD")
  NAME "ARTrans.Dat" AS "ARTrans.OLD"

  ARFile = FREEFILE
  OPEN "ARCUST.Dat" FOR RANDOM AS #ARFile LEN = LEN(ARCust(1))
  NumRecs = LOF(ARFile) / LEN(ARCust(1))
  
  FOR Cnt& = 1 TO NumRecs
    GET ARFile, Cnt&, ARCust(1)
    ARCust(1).FirstTrans = 0
    ARCust(1).LastTrans = 0
    ARCust(1).AcctBal = 0
    PUT ARFile, Cnt&, ARCust(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KillFile("FIXMANTO.EXE")

