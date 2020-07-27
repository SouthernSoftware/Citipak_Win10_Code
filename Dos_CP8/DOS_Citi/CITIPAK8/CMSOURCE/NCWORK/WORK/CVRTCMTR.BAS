DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  cr$ = CHR$(1)

  CLS


  PRINT "Processing Transaction. . ."
  PRINT
  PRINT
  
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS 1 LEN = CMTRRecLen
  NumRT& = LOF(1) / CMTRRecLen

  FOR ZZ& = 1 TO NumRT&
'    LOCATE 7, 1
    GET 1, ZZ&, CMTRRec(1)
    IF CMTRRec(1).ChkByte = cr$ THEN
      PRINT "Data allready converted. . ."
      GOTO DoneAllReady
    END IF
    CMTRRec(1).TransVoidNum = 0
    CMTRRec(1).ChkByte = cr$
    IF (CMTRRec(1).TransCash <> 0) AND (CMTRRec(1).TransCheck <> 0) THEN
      CMTRRec(1).TransTender = 3
    ELSEIF CMTRRec(1).TransCash <> 0 THEN
      CMTRRec(1).TransTender = 1
    ELSEIF CMTRRec(1).TransCheck <> 0 THEN
      CMTRRec(1).TransTender = 2
    ELSE
      CMTRRec(1).TransTender = 0
    END IF
    PUT #1, ZZ&, CMTRRec(1)
  NEXT

DoneAllReady:
  CLOSE

  CALL KillFile("cvrtcmtr.exe")
  PRINT "Done."

