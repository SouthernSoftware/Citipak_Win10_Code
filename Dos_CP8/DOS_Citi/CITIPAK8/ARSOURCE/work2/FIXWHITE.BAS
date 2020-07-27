DEFINT A-Z

  ''$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS

  DIM ARCustRec(1) AS ARCustRecType

  ARCustRecLen = LEN(ARCustRec(1))

  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

  FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, ARCustRec(1)
    SELECT CASE Cnt
    'CASE 82
    'STOP
    CASE 29, 60, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84
      ARCustRec(1).AcctBal = 0
      PUT ARFile, Cnt, ARCustRec(1)
    CASE ELSE
    END SELECT
  NEXT
  CLOSE
  CALL KillFile("FIXWHITE.EXE")

