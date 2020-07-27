  DEFINT A-Z
  
  '$INCLUDE: 'Newcust.bi'
  CLS
  LOCATE 3, 1
  PRINT "Clearing Est. Flag."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  OPEN "ubcust.dat" FOR RANDOM AS #1 LEN = UBCustRecLen
  NumCust = LOF(1) \ UBCustRecLen

  FOR Cnt = 1 TO NumCust
    LOCATE 6, 1: PRINT "Processing:"; Cnt; " of"; NumCust;
    GET #1, Cnt, UBCustRec(1)
    UBCustRec(1).EstFlag = ""
    PUT #1, Cnt, UBCustRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete."

