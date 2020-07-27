DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'
  ''$INCLUDE: 'fieldinf.BI'

CONST BlankInt% = -32767                        'Specifies blank integer
CONST BlankLng& = -2147483647                   '       "        long int.
CONST BlankSng! = -3.402823E+38                 '       "        single
CONST BlankDbl# = -1.79769313486231D+308        'Specifies blank double

  CLS
  PRINT
  PRINT
  PRINT "Processing Flat Rates. . ."
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
    DidEm = 0
    FOR FCnt = 1 TO 4
      FRCode$ = QPTrim$(UBCustRec(1).FlatRates(FCnt).FRDESC)
      IF FRCode$ = "0" THEN
        DidEm = -1
        UBCustRec(1).FlatRates(FCnt).FRDESC = ""
        UBCustRec(1).FlatRates(FCnt).FRAMT = BlankDbl#
        UBCustRec(1).FlatRates(FCnt).FRFREQ = ""
        UBCustRec(1).FlatRates(FCnt).REVSRC = BlankInt%
        UBCustRec(1).FlatRates(FCnt).NumMin = BlankInt%
      END IF
    NEXT
    IF DidEm THEN
      FixedCnt = FixedCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; FixedCnt

