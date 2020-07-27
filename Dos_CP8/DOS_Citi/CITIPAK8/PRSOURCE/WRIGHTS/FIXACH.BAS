DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRFiles.bi'

  CLS
  PRINT "Correcting Draft Code. . . "
  PRINT
  PRINT
  REDIM Emp2(1) AS EmpData2Type
  Emp2Len = LEN(Emp2(1))

  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2Len

  NumOfRec = LOF(1) \ Emp2Len

  FOR Cnt = 1 TO NumOfRec
    
    GET #1, Cnt, Emp2(1)
    Bank$ = QPTrim$(Emp2(1).BANKNAME)
    Emp2(1).PreNoteFlag = 0
    'IF LEN(Bank$) > 0 THEN
    '  Fixed = Fixed + 1
    '  Emp2(1).DRAFTCOD = "Y"
      PUT #1, Cnt, Emp2(1)
    'END IF
  NEXT
  CLOSE
  PRINT "Corrected:"; Fixed

