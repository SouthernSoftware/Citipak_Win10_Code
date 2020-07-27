DEFINT A-Z

  '$INCLUDE: 'ap.bi'

  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))

  PRINT "Processing. . ."
  APFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM AS APFile LEN = LdRecLen
  NumTrans& = LOF(APFile) \ LdRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET APFile, Cnt&, APLedgerRec(1)
    APLedgerRec(1).PSLFlag = "Y"
    PUT APFile, Cnt&, APLedgerRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL killfile("setpsl.exe")

