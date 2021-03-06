DEFINT A-Z
DECLARE FUNCTION Qptrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Converting Cubic Meters. . ."
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
    IF UBCustRec(1).Status = "A" AND UBCustRec(1).DelFlag = 0 THEN

      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).NumUser > 1 THEN
           Rate$ = Qptrim$(UBCustRec(1).Serv(1).RateCode)
           IF Rate$ = "01" THEN
             FoundCnt = FoundCnt + 1
           END IF
        END IF
      NEXT
    END IF
    'PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Processing Complete."; FoundCnt

