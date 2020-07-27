DEFINT A-Z

'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PRW2.bi'


CONST False = 0
CONST True = NOT False
  CLS
  PRINT "Correcting W2 file. . ."


  REDIM W2InfoRec(1)   AS W2FormType
  W2InfoRecLen = LEN(W2InfoRec(1))

  OPEN W2InfoFile FOR RANDOM AS #1 LEN = W2InfoRecLen
  NumOfRecs = LOF(1) \ W2InfoRecLen
  PUT #1, NumOfRecs + 1, W2InfoRec(1)
  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  PRINT
  END

