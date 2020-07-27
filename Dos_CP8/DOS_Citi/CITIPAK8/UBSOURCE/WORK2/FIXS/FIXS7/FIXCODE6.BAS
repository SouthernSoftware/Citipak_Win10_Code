DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Fixing the shit. . ."
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
    IF UBCustRec(1).Status = "A" THEN
      FOR zz = 1 TO 4
        IF UBCustRec(1).FlatRates(zz).RevSrc > 0 THEN
          IF UBCustRec(1).FlatRates(zz).FRAMT < 0 THEN
            STOP
          END IF
        END IF
      NEXT
    END IF

'    wCode$ = QPTrim$(UBCustRec(1).Serv(1).RateCode)
'    sCode$ = QPTrim$(UBCustRec(1).Serv(2).RateCode)
'    IF LEN(sCode$) > 0 THEN
'      UBCustRec(1).Serv(2).RateCode = "SW1"
'      Fixed = Fixed + 1
'      PUT UBFile, Cnt&, UBCustRec(1)
'    END IF
'    IF LEN(wCode$) > 0 THEN
'      UBCustRec(1).Serv(1).RateCode = "WA1"
      'IF UBCustRec(1).Book = "02" AND UBCustRec(1).Status = "A" THEN
      '  Fixed = Fixed + 1
      'END IF
'      PUT UBFile, Cnt&, UBCustRec(1)
'    END IF

  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Corrected:"; Fixed

