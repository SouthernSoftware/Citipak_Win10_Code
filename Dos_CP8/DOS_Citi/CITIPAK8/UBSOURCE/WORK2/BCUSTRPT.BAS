DEFINT A-Z
DECLARE FUNCTION uRound# (DblNum#)

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
    'LOCATE , 1
    'PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).PrevBalance < 0 THEN
      'STOP
    END IF
    IF UBCustRec(1).CurrBalance < 0 THEN
      'LPRINT LEFT$(UBCustRec(1).CustName, 25), UBCustRec(1).CurrBalance, UBCustRec(1).PrevBalance, UBCustRec(1).Status
      'STOP
    END IF
    'IF UBCustRec(1).Status = "B" THEN
      'IF UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance > 0 THEN
      '  Tot# = uRound(Tot# + UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
        'LPRINT Cnt&, UBCustRec(1).CustName, USING "####.##"; UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance
      'END IF
    'END IF
  NEXT
  CLOSE UBFile
  'LPRINT
  'LPRINT "Total:"; USING "#####.##"; Tot#
'PRINT
'PRINT
'PRINT
'PRINT "Done."

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

