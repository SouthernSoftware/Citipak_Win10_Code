DEFINT A-Z
DECLARE FUNCTION Date2Num (thedate$)
  '$INCLUDE: 'ARV95.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'


  NewDate = Date2Num("06-30-2002")

  DIM SHARED ARCustRec(1) AS ARCustRecType
  RecLen = LEN(ARCustRec(1))

  OPEN "arcust.dat" FOR RANDOM AS #1 LEN = RecLen
  FOR zz = 1 TO LOF(1) / RecLen
    GET #1, zz, ARCustRec(1)
    ARCustRec(1).VALID = NewDate
    PUT #1, zz, ARCustRec(1)
  NEXT
  CLOSE
  CALL KillFile("fixedate.exe")

