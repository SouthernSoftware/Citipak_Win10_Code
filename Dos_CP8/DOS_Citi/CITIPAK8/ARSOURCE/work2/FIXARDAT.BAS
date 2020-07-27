DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  BDate1 = Date2Num("01-01-1900")
  BDate2 = Date2Num("12-31-1900")
  BDate3 = Date2Num("01-01-1901")
  BDate4 = Date2Num("12-31-1901")
  BDate5 = Date2Num("01-01-1902")
  BDate6 = Date2Num("12-31-1902")
  
  '$INCLUDE: 'ARV90.bi'                        'A/R FILE LAYOUTS

  REDIM ARTransRec(1) AS ARTransRecType

  ARTRFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #ARTRFile LEN = LEN(ARTransRec(1))
  TrNumRecs = LOF(ARTRFile) / LEN(ARTransRec(1))

  FOR Cnt& = 1 TO TrNumRecs
    GET ARTRFile, Cnt&, ARTransRec(1)
    TRDate = ARTransRec(1).TransDate
    SELECT CASE TRDate
    CASE BDate1 TO BDate2
      Fixed = Fixed + 1
      BadDate$ = LEFT$(Num2Date(TRDate), 6) + "2000"
      ARTransRec(1).TransDate = Date2Num(BadDate$)
      PUT ARTRFile, Cnt&, ARTransRec(1)

    CASE BDate3 TO BDate4
      Fixed = Fixed + 1
      BadDate$ = LEFT$(Num2Date(TRDate), 6) + "2001"
      ARTransRec(1).TransDate = Date2Num(BadDate$)
      PUT ARTRFile, Cnt&, ARTransRec(1)

    CASE BDate5 TO BDate6
      Fixed = Fixed + 1
      BadDate$ = LEFT$(Num2Date(TRDate), 6) + "2002"
      ARTransRec(1).TransDate = Date2Num(BadDate$)
      PUT ARTRFile, Cnt&, ARTransRec(1)
    CASE ELSE
    END SELECT
  NEXT
  CLOSE
  PRINT "Fixed:"; Fixed

  CALL KILLFile("fixarDAT.exe")

