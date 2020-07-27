DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  CLS
  PRINT
  PRINT
  PRINT "Moving Meters. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Serv(2).RMTRTYPE = "S" THEN
      IF LEN(QPTrim$(UBCustRec(1).Serv(1).RMTRTYPE)) = 0 THEN
        SWAP UBCustRec(1).LocMeters(2), UBCustRec(1).LocMeters(1)
        'PRINT UBCustRec(1).CustName
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
    END IF
  NEXT
  CLOSE UBFile
  
  PRINT
  PRINT
  PRINT "Done."
  PRINT
  PRINT

