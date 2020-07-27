DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'cmfiles.BI'

  CLS
  PRINT "Processing Transactions"
  PRINT
  PRINT

  ADate = Date2Num("07-25-2001")

  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))

  REDIM UBTransRec(1)  AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

'
  OPEN "UBTRANS.DAT" FOR RANDOM AS #1 LEN = UBTransRecLen
  NumTrans& = LOF(1) \ UBTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, UBTransRec(1)
    IF UBTransRec(1).TransDate <= ADate THEN
      UBTransRec(1).Posted2GL = "Y"
        PUT #1, Cnt&, UBTransRec(1)
  '      Found& = Found& + 1
      END IF
  NEXT
  CLOSE

  OPEN "CMTRANS.DAT" FOR RANDOM AS #1 LEN = CMTransRecLen
  NumTrans& = LOF(1) \ CMTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, CMTransRec(1)
    IF CMTransRec(1).TransDate <= ADate THEN
      CMTransRec(1).Trans2GL = "Y"
      PUT #1, Cnt&, CMTransRec(1)
  '      Found& = Found& + 1
      END IF
  NEXT
  CLOSE

  PRINT "Done"



  CALL KILLFile("SETGRAB.EXE")

  END

