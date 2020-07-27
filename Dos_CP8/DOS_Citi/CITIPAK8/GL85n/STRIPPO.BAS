DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'ap.bi'                        'GL File Types
  CLS
  PRINT "Removing PO's. . ."
  DIM POTrans AS GLTransRecType
  TransRecLen = LEN(POTrans)


  CALL KillFile("potrans.old")

  NAME "potrans.dat" AS "potrans.old"

  OPEN "POTRANS.old" FOR RANDOM SHARED AS #1 LEN = TransRecLen
  OPEN "POTRANS.dat" FOR RANDOM SHARED AS #2 LEN = TransRecLen

  NumTrans& = LOF(1) \ TransRecLen
  
  FOR TRRec& = 1 TO NumTrans&
    LOCATE 3, 1
    PRINT "Processing:"; TRRec&; "of"; NumTrans&;
    GET #1, TRRec&, POTrans

    SELECT CASE TRRec&
    CASE 14864, 14865, 14932, 14933, 14934, 14935, 15078, 15079, 16930, 20828
  '    OCnt = OCnt + 1
    CASE ELSE
      PUT #2, , POTrans
    END SELECT
  NEXT
  CLOSE

'  STOP

  DIM APLedger AS APLedger81RecType
  APLRecLen = LEN(APLedger)

  CALL KillFile("apledger.old")
  NAME "apledger.dat" AS "apledger.old"

  OPEN "APLEDGER.OLD" FOR RANDOM SHARED AS #1 LEN = APLRecLen
  OPEN "APLEDGER.dat" FOR RANDOM SHARED AS #2 LEN = APLRecLen
  NumTrans& = LOF(1) \ APLRecLen

  FOR TRRec& = 1 TO NumTrans&
    LOCATE 5, 1
    PRINT "Processing:"; TRRec&; "of"; NumTrans&;
    GET #1, TRRec&, APLedger
    PONUM$ = QPTrim$(APLedger.PONUM)
    SELECT CASE PONUM$
    CASE "3410", "3435", "3462", "3806", "4545"
      APCnt = APCnt + 1
    CASE ELSE
      PUT #2, , APLedger
    END SELECT
  NEXT
  CLOSE


  LOCATE 8, 1
  PRINT "Processing Complete."

  CALL KillFile("strippo.exe")

