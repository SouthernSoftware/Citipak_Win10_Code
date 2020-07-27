DEFINT A-Z
'$INCLUDE: 'gl.BI'

CLS
PRINT "Correcting G/L Accounts. . ."
DIM GLRec AS GLAcctRecType              'Account Record Type: ? bytes

GLRecLen = LEN(GLRec)

  OPEN "glacct.dat" FOR RANDOM AS #1 LEN = GLRecLen
  NumOfRec& = LOF(1) / GLRecLen

  FOR zz& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; zz&; " of "; NumOfRec&
    GET #1, zz&, GLRec
    GLRec.BegBal = 0
    PUT #1, zz&, GLRec
  NEXT
  CLOSE
  LOCATE 12, 1
  PRINT "Done."

