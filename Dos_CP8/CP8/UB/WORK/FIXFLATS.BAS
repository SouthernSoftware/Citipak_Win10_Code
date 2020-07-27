DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Flat Rates. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO NumCust&
    LOCATE 7, 1
    PRINT "Processing:"; CCnt&; "of"; NumCust&;
    DidEm = 0
    GET UBCFile, CCnt&, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      FOR ZZc = 1 TO 4
        zz$ = STR$(UBCustRec(1).FlatRates(ZZc).FRAmt)
        IF INSTR(zz$, "QNAN") > 0 THEN
          UBCustRec(1).FlatRates(ZZc).FRAmt = 0
          DidEm = -1
        END IF
        IF UBCustRec(1).FlatRates(ZZc).FRAmt > 10000 THEN
          UBCustRec(1).FlatRates(ZZc).FRAmt = 0
          DidEm = -1
          'UBCustRec(1).FlatRates(ZZc).FRAmt = 0
          'UBCustRec(1).FlatRates(ZZc).RevSrc = 0
          'UBCustRec(1).FlatRates(ZZc).NumMin = 0
        END IF
      NEXT
      IF DidEm = -1 THEN
        DidCnt = DidCnt + 1
        PUT UBCFile, CCnt&, UBCustRec(1)
      END IF
    END IF
    
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Changed:"; DidCnt
PRINT
PRINT "Processing Complete."

CALL KillFile("fixflats.exe")

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

