DEFINT A-Z

'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PRW2.bi'


CONST False = 0
CONST True = NOT False
  CLS
  PRINT "Correcting Fed & State Gross wages. . ."

  FirstTime = True

  REDIM W2InfoRec(1)   AS W2FormType
  W2InfoRecLen = LEN(W2InfoRec(1))

  OPEN W2InfoFile FOR RANDOM AS #1 LEN = W2InfoRecLen
  NumOfRecs = LOF(1) \ W2InfoRecLen
  FOR cnt = 1 TO NumOfRecs
    GET #1, cnt, W2InfoRec(1)
    IF W2InfoRec(1).BOX13AMT > 0 THEN
      W2InfoRec(1).FEDWAGE = W2InfoRec(1).SOCWAGE - W2InfoRec(1).BOX13AMT
      W2InfoRec(1).STAWAGE = W2InfoRec(1).FEDWAGE
      Fixed = Fixed + 1
    ELSE
      W2InfoRec(1).FEDWAGE = W2InfoRec(1).SOCWAGE
      W2InfoRec(1).STAWAGE = W2InfoRec(1).FEDWAGE
    END IF
    PUT #1, cnt, W2InfoRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  PRINT
  END

