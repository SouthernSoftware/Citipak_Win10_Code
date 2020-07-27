DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  BDate1 = Date2Num("12-18-2004")
  BDate2 = Date2Num("02-18-2004")
  
  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS
  CLS
  PRINT "Processing. . ."

  REDIM ARTransRec(1) AS ARTransRecType

  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))

  FOR Cnt& = 1 TO TrNumRecs
    GET ARTRFile, Cnt&, ARTransRec(1)
    TRDate = ARTransRec(1).TransDate
    IF TRDate = BDate1 THEN
      Fixed = Fixed + 1
      ARTransRec(1).TransDate = BDate2
      PUT ARTRFile, Cnt&, ARTransRec(1)
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT

  PRINT "Fixed:"; Fixed

  CALL KILLFile("fixatmor.exe")

