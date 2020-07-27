DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNum%)

  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS

  DIM ARTransRec(1) AS ARTransRecType

  CLS
  LOCATE 3, 1
  PRINT "Correcting Dates. . ."

  ArTrnLen = LEN(ARTransRec(1))

  ARTrnFile = FREEFILE
  OPEN "ARTrans.Dat" FOR RANDOM SHARED AS ARTrnFile LEN = ArTrnLen

  NumTrnRecs& = LOF(ARTrnFile) / ArTrnLen

  FOR Cnt& = 1 TO NumTrnRecs&
    GET ARTrnFile, Cnt&, ARTransRec(1)

    ThisDate$ = Num2Date(ARTransRec(1).TransDate)

    SELECT CASE RIGHT$(ThisDate$, 4)
    CASE "1900"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2000"
      GOSUB UPDateTrans
    CASE "1901"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2001"
      GOSUB UPDateTrans
    CASE "1902"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2002"
      GOSUB UPDateTrans
    CASE "1903"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2003"
      GOSUB UPDateTrans
    CASE "1904"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2004"
      GOSUB UPDateTrans
    CASE "1905"
      ThisDate$ = LEFT$(ThisDate$, 6) + "2005"
      GOSUB UPDateTrans
    CASE ELSE
    'don't do anything
    END SELECT

  NEXT
  CLOSE

  LOCATE 12, 4
  PRINT "Corrected:"; Fixed&
  PRINT
  CALL KillFile("FIXDATES.EXE")

END

UPDateTrans:
  Fixed& = Fixed& + 1
  ARTransRec(1).TransDate = Date2Num(ThisDate$)
  PUT ARTrnFile, Cnt&, ARTransRec(1)
RETURN

