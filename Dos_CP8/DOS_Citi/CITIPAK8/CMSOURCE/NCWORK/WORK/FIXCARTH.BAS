DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'CMFILES.BI'

  CLS
  PRINT "Processing Transactions. . ."
  CALL KillFile("CMTRANS.BAK")
  NAME "CMTRANS.DAT" AS "CMTRANS.BAK"

  BadDate = Date2Num("10-01-2010")

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  TrHandle = FREEFILE
  OPEN "CMTRANS.BAK" FOR RANDOM SHARED AS TrHandle LEN = CMTRRecLen
  TrNumRecs& = LOF(TrHandle) \ CMTRRecLen
  TrHandle2 = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS TrHandle2 LEN = CMTRRecLen
  FOR Cnt& = 1 TO TrNumRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; TrNumRecs&;
    GET TrHandle, Cnt&, CMTRRec(1)
    IF CMTRRec(1).TransDate <> BadDate THEN
      PUT TrHandle2, , CMTRRec(1)
    ELSE
      BadCount& = BadCount& + 1
    END IF
  NEXT
  CLOSE
  LOCATE 10, 1
  PRINT "Removed:"; BadCount&
  CALL KillFile("fixcarth.exe")





