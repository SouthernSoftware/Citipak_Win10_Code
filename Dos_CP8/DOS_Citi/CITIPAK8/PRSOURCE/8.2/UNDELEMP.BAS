DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRFiles.bi'
  CLS
  PRINT "Correcting Deleted Employee's"
  REDIM TEmp2Rec(1) AS EmpData2Type
  Emp2RecLen = LEN(TEmp2Rec(1))

  REDIM TEmp1Rec(1) AS EmpData1Type
  Emp1RecLen = LEN(TEmp1Rec(1))

  OPEN EmpData1Name FOR RANDOM AS #2 LEN = Emp1RecLen
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2RecLen
  NumOfRec = LOF(2) \ Emp1RecLen

  FOR Cnt = 1 TO NumOfRec
    GET #2, Cnt, TEmp1Rec(1)
    GET #1, Cnt, TEmp2Rec(1)
    TEmp1Rec(1).Deleted = TEmp2Rec(1).Deleted
    PUT #2, Cnt, TEmp1Rec(1)
    EmpNumb = VAL(QPTrim$(TEmp1Rec(1).EMPNO))
    SELECT CASE EmpNumb
      CASE 112 TO 127, 132, 133, 140 TO 143, 152, 163 TO 166
        TEmp1Rec(1).Deleted = 0
        TEmp2Rec(1).Deleted = 0
        PUT #1, Cnt, TEmp2Rec(1)
        PUT #2, Cnt, TEmp1Rec(1)
      CASE ELSE
    END SELECT
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Done!"

