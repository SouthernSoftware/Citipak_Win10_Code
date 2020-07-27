DEFINT A-Z

'$INCLUDE: 'GL.BI'


SUB GetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)

  REDIM GLSetUpRec(1) AS GLSetupRecType
  SetUpRecLen = LEN(GLSetUpRec(1))

  SetupFile = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen
  GET SetupFile, 1, GLSetUpRec(1)

  GLFundLen = GLSetUpRec(1).FundLen
  GLAcctLen = GLSetUpRec(1).AcctLen
  GLDetLen = GLSetUpRec(1).DetLen

  CLOSE SetupFile
  ERASE GLSetUpRec
END SUB

