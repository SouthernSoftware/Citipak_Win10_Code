DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'

  CLS
  PRINT "Removing Tax Transaction History. . ."
  NAME "taxtrans.dat" AS "txtran.old"

  REDIM TaxCustRec(1) AS TaxCustType
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxFile = FREEFILE
  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen
  NumOfCust& = LOF(TaxFile) / TaxCustRecLen
  FOR Cnt& = 1 TO NumOfCust&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; " of"; NumOfCust&;
    GET TaxFile, Cnt&, TaxCustRec(1)
    TaxCustRec(1).LastTrans = 0
    PUT TaxFile, Cnt&, TaxCustRec(1)
  NEXT
  CLOSE
  LOCATE 8, 1
  PRINT "Processing Complete."
  KILL "CLEARHST.EXE"

