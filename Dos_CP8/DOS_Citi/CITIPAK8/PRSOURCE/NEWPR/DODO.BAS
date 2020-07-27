DEFINT A-Z

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRFiles.bi'

'EDist(1 TO 8)   AS EmpWageDistType
DIM Emp2 AS EmpData2Type
E2Len = LEN(Emp2)

OPEN EmpData2Name FOR RANDOM AS #1 LEN = E2Len
FOR Cnt = 1 TO LOF(1) \ E2Len
  GET #1, Cnt, Emp2
  LPRINT Emp2.EMPLNAME, Emp2.EMPFNAME

  FOR zz = 1 TO 8
    IF Emp2.EDist(zz).DAmt > 0 THEN
      LPRINT Emp2.EDist(zz).DAcct, Emp2.EDist(zz).DAmt
    END IF
  NEXT
NEXT
CLOSE

