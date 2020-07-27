DEFINT A-Z

  '$INCLUDE: 'ap.bi'

  CONST False = 0, True = NOT False, NumAcctFlds = 24
  
  REDIM APLedgerRec(1) AS APLedger81RecType
  LdRecLen = LEN(APLedgerRec(1))

  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM SHARED AS APLedgerFile LEN = LdRecLen
  NumTrans = LOF(APLedgerFile) \ LdRecLen
  FOR Cnt& = 1 TO NumTrans
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumTrans;
    GET APLedgerFile, Cnt&, APLedgerRec(1)
    APLedgerRec(1).PSLFlag = "Y"
    PUT APLedgerFile, Cnt&, APLedgerRec(1)
  NEXT
  CLOSE

