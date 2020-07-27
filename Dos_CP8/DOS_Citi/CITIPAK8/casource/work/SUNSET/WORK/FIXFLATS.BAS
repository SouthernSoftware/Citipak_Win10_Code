DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    LOCATE 6, 1
    PRINT "Processing:"; Cnt&; "of"; CNumOfRecs&;
    GET UBCFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Book <> "04" THEN
      FOR FRCnt = 1 TO 4
        UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
        UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
        UBCustRec(1).FlatRates(FRCnt).NumMin = 0
        UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
        UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
      NEXT
      Fixed = Fixed + 1
    END IF
    PUT UBCFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  PRINT "Fixed:"; Fixed

  KillFile "FIXFLATS.EXE"


