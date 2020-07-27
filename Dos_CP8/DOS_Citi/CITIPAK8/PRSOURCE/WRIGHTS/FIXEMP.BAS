DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRFiles.bi'

  REDIM TEmp2Rec(1) AS EmpData2Type
  Emp2RecLen = LEN(TEmp2Rec(1))
  CLS
  PRINT "Processing. . ."
  PRINT
  PRINT
  OPEN EmpData2Name FOR RANDOM AS #1 LEN = Emp2RecLen

  NumOfRec = LOF(1) \ Emp2RecLen

  FOR Cnt = 1 TO NumOfRec

    GET #1, Cnt, TEmp2Rec(1)
    EmpNumb$ = TEmp2Rec(1).EMPNO

    IF INSTR(EmpNumb$, "5100029") OR INSTR(EmpNumb$, "5100030") THEN
      TEmp2Rec(1).PreNoteFlag = 0
      Fixed = Fixed + 1
      PUT #1, Cnt, TEmp2Rec(1)
    END IF
  NEXT
  CLOSE

  PRINT "Corrected:"; Fixed

