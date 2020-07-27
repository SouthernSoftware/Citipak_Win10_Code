  DEFINT A-Z
  '$INCLUDE: 'GL.BI'                        'GL File Types
  
  DIM Acct AS GLAcctRecType
  AcctLen = LEN(Acct)
  OPEN "glacct.DAT" FOR RANDOM AS #1 LEN = AcctLen
  FOR zz = 1 TO LOF(1) / AcctLen
    GET #1, zz, Acct
    Acct.FrstTran = 0
    Acct.LastTran = 0
    Acct.PYAct = 0
    Acct.BegBal = 0
    Acct.Bgt = 0
    Acct.Bal = 0
    Acct.Encumb = 0
    Acct.MTD = 0
    Acct.YTD = 0
    Acct.NYEst = 0
    Acct.NYReq = 0
    Acct.NYRec = 0
    Acct.NYApp = 0
    Acct.FrstBTran = 0
    Acct.LastBTran = 0
    Acct.FrstPTran = 0
    Acct.LastPTran = 0
    PUT #1, zz, Acct
  NEXT
  CLOSE
  CALL KILLFILE("ZEROBAL.EXE")

'  DIM BgtTrans AS GLTransRecType
'  BgtTransFile = FREEFILE
'  BgtTransRecLen = LEN(BgtTrans)
'  OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen
'  NumBgtTrans = LOF(BgtTransFile) \ BgtTransRecLen
'
'
'
'  FOR Cnt = 1 TO NumBgtTrans
'    GET #BgtTransFile, Cnt, BgtTrans
'    '     IF BgtTrans.CrAmt > 0 OR BgtTrans.DrAmt > 0 THEN STOP
'    BgtTrans.CrAmt = 0
'    BgtTrans.DrAmt = 0
'    PUT #BgtTransFile, Cnt, BgtTrans
'  NEXT
'
'  CLOSE
  

