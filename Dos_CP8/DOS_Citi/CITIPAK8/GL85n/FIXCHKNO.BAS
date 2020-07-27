DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'AP.bi'                        'AP File Types

  CLS
  PRINT "Correcting Check Numbers. . ."
  TDate = Date2Num("11-01-2001")

  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)

  APLedgerFile = FREEFILE
  OPEN "APLEDGER.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS APLedgerFile LEN = APLRecLen
  NumTrans& = LOF(APLedgerFile) \ APLRecLen
  FOR Cnt& = 1 TO NumTrans&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumTrans&;
    GET APLedgerFile, Cnt&, APLedger
    IF APLedger.TrCode = 3 OR APLedger.TrCode = -3 THEN
      IF APLedger.TrDate = TDate THEN
        Fixed = Fixed + 1
        Chk& = VAL(QPTrim$(APLedger.DocNum))
        Chk& = Chk& + 1
        APLedger.DocNum = QPTrim$(STR$(Chk&))
        PUT APLedgerFile, Cnt&, APLedger
      END IF
    END IF
  NEXT
  CLOSE
  LOCATE 10, 1
  PRINT "Fixed:"; Fixed
  CALL killfile("fixchkno.exe")

