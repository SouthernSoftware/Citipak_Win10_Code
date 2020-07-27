DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  CONST False = 0, True = NOT False

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  LOCATE 2, 1, 0
  PRINT "Processing Accounts. . ."

  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = False
    FOR Mtr = 1 TO 7
      AvgUse# = UBCustRec(1).LocMeters(Mtr).AvgUse
      IF AvgUse# < 0 THEN
        AvgUse# = 0
      END IF
      UBCustRec(1).MeterAux(Mtr).MaxUsed = AvgUse#
      UBCustRec(1).MeterAux(Mtr).MinUsed = AvgUse#
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Complete."
  PRINT
  PRINT

