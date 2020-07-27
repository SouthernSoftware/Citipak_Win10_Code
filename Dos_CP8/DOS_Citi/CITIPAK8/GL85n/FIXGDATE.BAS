DEFINT A-Z
DECLARE FUNCTION Date2Num% (d$)
DECLARE FUNCTION Num2Date$ (DNum%)

'$INCLUDE: 'gl.BI'

  REDIM GLRec(1) AS GLTransRecType
  GLRecLen = LEN(GLRec(1))

  FunkyDate = Date2Num("05-02-2003")
  GoodDate = Date2Num("01-25-2003")

  CLS
  PRINT "Adjusting Dates. . ."
  OPEN "gltrans.dat" FOR RANDOM AS #1 LEN = GLRecLen
  NumOfRec& = LOF(1) \ GLRecLen

  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 3, 1: PRINT "Examining:"; Cnt&; " of"; NumOfRec&;
    GET #1, Cnt&, GLRec(1)
    IF GLRec(1).TrDate = FunkyDate THEN
      DidCnt = DidCnt + 1
      GLRec(1).TrDate = GoodDate
      PUT #1, Cnt&, GLRec(1)
    END IF
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Did:"; DidCnt
  PRINT "Processing Complete."
  CALL KillFile("fixgdate.exe")

