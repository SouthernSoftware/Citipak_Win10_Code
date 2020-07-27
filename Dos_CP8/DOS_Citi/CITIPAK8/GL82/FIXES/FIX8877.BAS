DEFINT A-Z

DECLARE FUNCTION Date2Num% (d$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'GL.bi'                        'GL File & Record Types

  CONST False = 0, True = NOT False

  COLOR 15, 1
  CLS

  PRINT "Correcting Transaction #8877"
  PRINT
  
  REDIM GLTrans(1) AS GLTransRecType
  GLRecLen = LEN(GLTrans(1))

  OPEN "GLTRANS.DAT" FOR RANDOM AS #1 LEN = GLRecLen
  GET #1, 8877, GLTrans(1)

  Acct$ = QPTrim$(GLTrans(1).AcctNum)   'if this is the wrong account
  IF Acct$ <> "30-101-01" THEN
    ErrFlag = True
  END IF

  TDate = Date2Num("04-17-1997")
  IF GLTrans(1).TrDate <> TDate THEN    'if the trans date is wrong
    ErrFlag = True
  END IF

  Desc$ = UCASE$(QPTrim$(GLTrans(1).Desc))
  IF Desc$ <> "CLEVELAND ANDREWS" THEN  'if the description is invalid
    ErrFlag = True
  END IF

  Source$ = QPTrim$(GLTrans(1).Src)
  IF Source$ <> "CD050697" THEN         'if the source is invalid
    ErrFlag = True
  END IF

  IF ErrFlag THEN                       'if there was an error
    PRINT "AN ERROR HAS OCCURED, PROCESSING THE INVALID TRANSACTION."
    PRINT "NO DATA HAS BEEN CHANGED!!!  TAKE NO OTHER ACTION!!!"
    PRINT "CALL SOUTHERN SOFTWARE NOW!!  1-800-842-8190"
    PRINT
  ELSE      'all is well
    GLTrans(1).DrAmt = 0
    GLTrans(1).CrAmt = 0
    PUT #1, 8877, GLTrans(1)
    PRINT "Transaction #8877 has been corrected."
    PRINT
    PRINT "Processing complete."
    PRINT
  END IF
  CLOSE
  END

