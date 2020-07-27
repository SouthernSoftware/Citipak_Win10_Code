DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  BDate1 = Date2Num("01-01-1901")
  BDate2 = Date2Num("12-31-1901")
  
  '$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS

  REDIM ARTransRec(1) AS ARTransRecType

  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))

  FOR Cnt& = 1 TO TrNumRecs
    GET ARTRFile, Cnt&, ARTransRec(1)
    TRDate = ARTransRec(1).TransDate
    IF TRDate >= BDate1 AND TRDate <= BDate2 THEN
      Fixed = Fixed + 1
      BadDate$ = LEFT$(Num2Date(TRDate), 6) + "2001"
      ARTransRec(1).TransDate = Date2Num(BadDate$)
      PUT ARTRFile, Cnt&, ARTransRec(1)
    END IF
  NEXT
  CLOSE
  PRINT "Fixed:"; Fixed

