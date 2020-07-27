  DEFINT A-Z
  '$INCLUDE: 'GL.BI'                        'GL File Types
  
  
  'DIM SHARED BgtEdit AS TrEditRecType
  '
  'BgtEdLen = LEN(BgtEdit)
  'BgtEditFileNum = FREEFILE
  'OPEN "BGTED.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BgtEditFileNum LEN = BgtEdLen
  'NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen
  
  DIM BgtTrans AS GLTransRecType
  BgtTransFile = FREEFILE
  BgtTransRecLen = LEN(BgtTrans)
  OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen
  NumBgtTrans = LOF(BgtTransFile) \ BgtTransRecLen
  
  
  
  FOR Cnt = 1 TO NumBgtTrans
    GET #BgtTransFile, Cnt, BgtTrans
    '     IF BgtTrans.CrAmt > 0 OR BgtTrans.DrAmt > 0 THEN STOP
    BgtTrans.CrAmt = 0
    BgtTrans.DrAmt = 0
    PUT #BgtTransFile, Cnt, BgtTrans
  NEXT
  
  CLOSE
  
