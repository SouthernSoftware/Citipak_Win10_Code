DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TDate1 = Date2Num%("01-13-1998")
  TDate2 = Date2Num%("02-11-1998")
  TDate3 = Date2Num%("02-23-1998")
  TDate4 = Date2Num%("01-05-1998")
  TDate5 = Date2Num%("01-27-1998")

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transactions. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  GET UBFile, 48500, UBTranRec(1)
  UBTranRec(1).TransDate = TDate1
  PUT UBFile, 48500, UBTranRec(1)

  GET UBFile, 48499, UBTranRec(1)
  UBTranRec(1).TransDate = TDate1
  PUT UBFile, 48499, UBTranRec(1)

  GET UBFile, 48498, UBTranRec(1)
  UBTranRec(1).TransDate = TDate1
  PUT UBFile, 48498, UBTranRec(1)

  GET UBFile, 53334, UBTranRec(1)
  UBTranRec(1).TransDate = TDate2
  PUT UBFile, 53334, UBTranRec(1)

  GET UBFile, 53335, UBTranRec(1)
  UBTranRec(1).TransDate = TDate2
  PUT UBFile, 53335, UBTranRec(1)

  GET UBFile, 53336, UBTranRec(1)
  UBTranRec(1).TransDate = TDate2
  PUT UBFile, 53336, UBTranRec(1)

  GET UBFile, 54021, UBTranRec(1)
  UBTranRec(1).TransDate = TDate3
  PUT UBFile, 54021, UBTranRec(1)

  GET UBFile, 47319, UBTranRec(1)
  UBTranRec(1).TransDate = TDate4
  PUT UBFile, 47319, UBTranRec(1)

  GET UBFile, 49470, UBTranRec(1)
  UBTranRec(1).TransDate = TDate5
  PUT UBFile, 49470, UBTranRec(1)

  GET UBFile, 17576, UBTranRec(1)
  UBTranRec(1).TransAmt = 94.69
  UBTranRec(1).RevAmt(1) = 28.77
  UBTranRec(1).RevAmt(2) = 53.92
  UBTranRec(1).RevAmt(3) = 12
  PUT UBFile, 17576, UBTranRec(1)

  GET UBFile, 18586, UBTranRec(1)
  UBTranRec(1).TransAmt = 60.72
  UBTranRec(1).RevAmt(1) = 19.52
  UBTranRec(1).RevAmt(2) = 33.2
  UBTranRec(1).RevAmt(3) = 8
  PUT UBFile, 18586, UBTranRec(1)

  GET UBFile, 19033, UBTranRec(1)
  UBTranRec(1).TransAmt = 48.34
  UBTranRec(1).RevAmt(1) = 8.56
  UBTranRec(1).RevAmt(2) = 31.78
  UBTranRec(1).RevAmt(3) = 8
  PUT UBFile, 19033, UBTranRec(1)

  CLOSE

LOCATE 9, 1
PRINT "  Done."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

