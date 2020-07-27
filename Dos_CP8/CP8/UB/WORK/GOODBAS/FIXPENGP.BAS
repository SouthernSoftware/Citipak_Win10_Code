DEFINT A-Z

'DECLARE FUNCTION FileSize& (FileName$)
'DECLARE FUNCTION Num2Date$ (DateNum%)
'DECLARE FUNCTION Date2Num% (Today$)
'DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'ubtrans.BI'
'  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False


  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Transaction. . ."

  PRINT
  PRINT

  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  'LOCATE 5, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
  Cnt& = 89132
  GET UBTFile, Cnt&, UBTranRec(1)
  UBTranRec(1).TransAmt = 4.5

  FOR zz = 1 TO 15
    UBTranRec(1).RevAmt(zz) = 0
    UBTranRec(1).TaxAmt(zz) = 0
  NEXT
  UBTranRec(1).RevAmt(3) = UBTranRec(1).TransAmt
  PUT UBTFile, Cnt&, UBTranRec(1)
  CLOSE

LOCATE 9, 1
PRINT "Done"

  CALL KillFile("FIXPENGP.EXE")


