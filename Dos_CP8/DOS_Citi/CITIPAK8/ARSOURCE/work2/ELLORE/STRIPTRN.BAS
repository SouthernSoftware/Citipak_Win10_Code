DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  BDate1 = Date2Num("07-06-1900")
  
  '$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  CLS
  PRINT "Processing. . ."

  REDIM ARTransRec(1) AS ARTransRecType

  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))

  ARTRFile2 = FREEFILE
  OPEN "ARTrans2.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile2 LEN = LEN(ARTransRec(1))

  FOR Cnt& = 1 TO TrNumRecs
    GET ARTRFile, Cnt&, ARTransRec(1)
    TRDate = ARTransRec(1).TransDate
    IF TRDate = BDate1 THEN
      Fixed = Fixed + 1
    ELSE
      'BadDate$ = LEFT$(Num2Date(TRDate), 6) + "2000"
      'ARTransRec(1).TransDate = Date2Num(BadDate$)
      PUT ARTRFile2, , ARTransRec(1)
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Corrected:"; Fixed


