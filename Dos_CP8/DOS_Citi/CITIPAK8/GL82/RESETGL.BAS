DEFINT A-Z

  '$INCLUDE: 'GL.BI'                        'GL File Types
  '$INCLUDE: 'ap.bi'

  CLS
  PRINT
  PRINT
  PRINT "Clearing History. . ."

  CALL KillFile("GLTRANS.DAT")
  CALL KillFile("APLEDGER.DAT")
  CALL KillFile("APDIST.DAT")
  CALL KillFile("BGTTRANS.DAT")

  DIM GLAcct  AS GLAcctRecType
  GLAcctLen = LEN(GLAcct)

  OPEN "glacct.dat" FOR RANDOM AS #1 LEN = GLAcctLen
  NumAcct = LOF(1) / GLAcctLen
  FOR Cnt = 1 TO NumAcct
    GET #1, Cnt, GLAcct
    GLAcct.FrstTran = 0            'Pointer to First Trans
    GLAcct.LastTran = 0            'Pointer to Last Trans
    GLAcct.PYAct = 0               'Prior Year Actual
    GLAcct.BegBal = 0              'Beginning Balance
    GLAcct.Bgt = 0                 'Budget Amount
    GLAcct.Bal = 0                 'Running Balance
    GLAcct.Encumb = 0              'Encumbered Amount
    GLAcct.MTD = 0                 'Month to Date Bal (calc as needed)
    GLAcct.YTD = 0                 'Year to Date Bal (calc as needed)
    GLAcct.NYEst = 0               'Bgt Estimate
    GLAcct.NYReq = 0               'New Year Bgt Requested
    GLAcct.NYRec = 0               'New Year Bgt Recommended
    GLAcct.NYApp = 0               'New Year Bgt Approved
    GLAcct.FrstBTran = 0           'Pointer to First Budget Trans
    GLAcct.LastBTran = 0           'Pointer to Last Budget Trans
    GLAcct.FrstPTran = 0           'Pointer to First Budget Trans
    GLAcct.LastPTran = 0           'Pointer to Last Budget Trans
    PUT #1, Cnt, GLAcct
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Processing Complete."

