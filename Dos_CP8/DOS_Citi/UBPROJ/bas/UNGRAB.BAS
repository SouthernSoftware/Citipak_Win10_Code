DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)

  '$INCLUDE: 'ubtrans.BI'

  CLS

  ADate = Date2Num("03-01-2001")
  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

'
  OPEN "UBTRANS.DAT" FOR RANDOM AS #1 LEN = UBTransRecLen
  NumTrans& = LOF(1) \ UBTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, UBTransRec(1)
    IF UBTransRec(1).TransDate >= ADate THEN
      IF UBTransRec(1).Posted2GL = "Y" THEN
        UBTransRec(1).Posted2GL = ""
        PUT #1, Cnt&, UBTransRec(1)
        Found = Found + 1
      END IF
    END IF
  NEXT
  CLOSE
  PRINT "UN-Grabbed:"; Found

  CALL KILLFile("UNGRAB.EXE")

  END

