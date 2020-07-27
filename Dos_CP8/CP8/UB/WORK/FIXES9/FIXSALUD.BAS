DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing User Fee. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumCust&
    GET UBCFile, Cnt&, UBCustRec(1)
    LOCATE 6, 1: PRINT "Processing:"; Cnt&; "of"; NumCust&;
    UBCustRec(1).Serv(5).RateCode = ""
    FOR RCnt = 1 TO 4
      IF UBCustRec(1).FlatRates(RCnt).REVSRC = 5 THEN
        UBCustRec(1).FlatRates(RCnt).FRDESC = ""
        UBCustRec(1).FlatRates(RCnt).FRAMT = 0
        UBCustRec(1).FlatRates(RCnt).FRFREQ = ""
        UBCustRec(1).FlatRates(RCnt).REVSRC = 0
        UBCustRec(1).FlatRates(RCnt).NumMin = 0
        Removed = Removed + 1
      END IF
    NEXT
    PUT UBCFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."
  LOCATE 12, 1
  PRINT "Removed:"; Removed
  KillFile "FIXSALUD.EXE"

