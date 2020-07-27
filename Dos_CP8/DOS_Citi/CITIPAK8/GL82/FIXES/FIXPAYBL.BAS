DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  '$INCLUDE: 'AP.BI'
  NeedDate = Date2Num("01-17-2001")
  FindDate = Date2Num("01-17-2000")

  REDIM APLedgerRec(1) AS APLedger81RecType
  RecLen = LEN(APLedgerRec(1))

  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APLedgerFile LEN = RecLen
  NumTran& = LOF(APLedgerFile) \ RecLen
  FOR Cnt& = 1 TO NumTran&
    GET APLedgerFile, Cnt&, APLedgerRec(1)
    IF APLedgerRec(1).TrCode = 1 AND APLedgerRec(1).PAYCODE = 1 THEN
      IF APLedgerRec(1).TRDate = FindDate AND APLedgerRec(1).DUEDATE = FindDate THEN
        APLedgerRec(1).TRDate = NeedDate
        APLedgerRec(1).DUEDATE = NeedDate
        PUT APLedgerFile, Cnt&, APLedgerRec(1)
        Fixed = Fixed + 1
      END IF
    END IF
  NEXT
  CLOSE
  PRINT "Fixed:"; Fixed
  CALL KillFile("fixpaybl.exe")

