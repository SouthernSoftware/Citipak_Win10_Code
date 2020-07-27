DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Monthly Billed Amounts. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT


  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust = LOF(UBFile) \ UBCustRecLen
  FOR Cnt = 1 TO NumOfCust
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; " of"; NumOfCust;
    GET #UBFile, Cnt, UBCustRec(1)
    FOR ZZ = 1 TO 2
      UBCustRec(1).Monthly(ZZ).AMTOWED = 0
      UBCustRec(1).Monthly(ZZ).TotAmtPD = 0
      UBCustRec(1).Monthly(ZZ).PayAmt = 0
      UBCustRec(1).Monthly(ZZ).RevSource = 0
    NEXT
    PUT #UBFile, Cnt, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 15, 1
  PRINT "Processing Complete."

