DEFINT A-Z

'$INCLUDE: 'GL.bi'                        'GL File Types
  CLS
  PRINT Processing; Transactions.

  DIM BgtTrans AS GLTransRecType
  BgtTransLen = LEN(BgtTrans)
  CALL killfile("BGTTRANS.bak")
  NAME "BGTTRANS.dat" AS "BGTTRANS.bak"

  BgtTransFile = FREEFILE
  OPEN "BGTTRANS.BAK" FOR RANDOM SHARED AS BgtTransFile LEN = BgtTransLen
  NumTrans = LOF(BgtTransFile) \ BgtTransLen

  BFile = FREEFILE
  OPEN "BGTTRANS.DAT" FOR RANDOM SHARED AS BFile LEN = BgtTransLen
  FOR Cnt& = 1 TO NumTrans
    GET BgtTransFile, Cnt&, BgtTrans
    SELECT CASE Cnt&
    CASE 232, 233, 234, 287
      'do nothing
    CASE ELSE
      PUT BFile, , BgtTrans
    END SELECT
  NEXT
  CLOSE

  LOCATE 7, 1
  PRINT "Processing Complete."
  CALL killfile("fixsnbch.exe")

