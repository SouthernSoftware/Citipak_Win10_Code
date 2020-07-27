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
  PRINT "Processing. . ."
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
      FOR zz = 1 TO 7
        MT$ = QPTrim$(UBCustRec(1).LocMeters(zz).MtrType)
        IF LEN(MT$) > 0 THEN
          IF UBCustRec(1).LocMeters(zz).MTRMulti = 100 THEN
            UBCustRec(1).LocMeters(zz).MTRMulti = 1000
            'CurRead& = UBCustRec(1).LocMeters(zz).CurRead
            'PreRead& = UBCustRec(1).LocMeters(zz).PrevRead
            'IF CurRead& > 1000 THEN
            '  CurRead& = CurRead& / 1000
            '  UBCustRec(1).LocMeters(zz).CurRead = CurRead&
            'END IF
            'IF PreRead& > 1000 THEN
            '  PreRead& = PreRead& / 1000
            '  UBCustRec(1).LocMeters(zz).PrevRead = PreRead&
            'END IF

            Found = Found + 1
            'EXIT FOR
          END IF
        END IF
      NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Found:"; Found
CALL KILLfile("fixmulti.exe")
    

