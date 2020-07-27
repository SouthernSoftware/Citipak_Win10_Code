DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  CLS
  Cmd$ = COMMAND$
  PRINT "FixTDate v1.1     Southern Software."
  PRINT

  IF LEN(Cmd$) < 21 THEN
    IF LEN(Cmd$) > 0 THEN
      PRINT "Invalid Date: "; Cmd$
      PRINT
    END IF
    PRINT " Syntax:  FIXTDATE {FromDate} {ToDate}"
    PRINT
    PRINT "   NOTE:  THE ENTIRE DATE MUST BE USED"
    PRINT
    PRINT "Example:  FIXTDATE 12-31-1979 01-31-1998"
    PRINT
    END
  END IF

  SPos = INSTR(Cmd$, " ")

  NDate = Date2Num%(LEFT$(Cmd$, SPos))
  TDate = Date2Num%(MID$(Cmd$, SPos + 1))

  LOCATE 3, 1, 0
  PRINT "Correcting Transaction Dates. . ."

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen
  'OPEN "TRANDATE.LST" FOR OUTPUT AS #10
  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
   ' PRINT #10, Num2Date(UBTranRec(1).TransDate)
    IF UBTranRec(1).TransDate = NDate THEN
      TTCnt = TTCnt + 1
      UBTranRec(1).TransDate = TDate
      PUT UBFile, Cnt&, UBTranRec(1)
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "  Fixed:"; TTCnt;

