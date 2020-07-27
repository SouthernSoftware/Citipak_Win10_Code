DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (FileName$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Processing Invalid Meter Settings. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF LEN(QPTrim$(UBCustRec(1).Serv(1).RateCode)) = 0 AND LEN(QPTrim$(UBCustRec(1).Serv(1).RMTRTYPE)) > 0 THEN
        UBCustRec(1).Serv(1).RMTRTYPE = ""
        PUT UBFile, Cnt&, UBCustRec(1)
        Found = Found + 1
      END IF
      IF LEN(QPTrim$(UBCustRec(1).Serv(2).RateCode)) = 0 AND LEN(QPTrim$(UBCustRec(1).Serv(2).RMTRTYPE)) > 0 THEN
        UBCustRec(1).Serv(2).RMTRTYPE = ""
        PUT UBFile, Cnt&, UBCustRec(1)
        Found = Found + 1
      END IF
    END IF
DidEm:
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Found:"; Found

END

PrintEM:
  TabStop = TabStop + 10
  IF TabStop > 70 THEN
    LPRINT
    TabStop = 10
  END IF
  LPRINT TAB(TabStop); Cnt&;

RETURN

