DEFINT A-Z
DECLARE FUNCTION Date2Num% (d$)
DECLARE FUNCTION Num2Date$ (DNum%)

'$INCLUDE: 'gl.BI'

REDIM GLRec(1) AS GLTransRecType
GLRecLen = LEN(GLRec(1))
CLS
PRINT "Adjusting Dates. . ."
OPEN "gltrans.dat" FOR RANDOM AS #1 LEN = GLRecLen
NumOfRec& = LOF(1) \ GLRecLen

FOR Cnt& = 1 TO NumOfRec&
  LOCATE 3, 1: PRINT "Examining:"; Cnt&; " of"; NumOfRec&;
  GET #1, Cnt&, GLRec(1)
  IF LEFT$(GLRec(1).Src, 2) = "AP" THEN
    APDate = Date2Num(MID$(GLRec(1).Src, 3))
    IF APDate <> GLRec(1).TrDate THEN
      DidCnt = DidCnt + 1
      GLRec(1).TrDate = APDate
      PUT #1, Cnt&, GLRec(1)
      'STOP
    END IF
  END IF
NEXT
CLOSE
LOCATE 8, 1
'PRINT "Did:"; DidCnt
PRINT "Processing Complete."

