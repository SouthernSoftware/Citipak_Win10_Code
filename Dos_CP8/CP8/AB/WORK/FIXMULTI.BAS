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
    'FOR zz = 1 TO 7
      UBCustRec(1).locmeters(1).MTRMulti = 1000
    'NEXT
    PUT UBCFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  PRINT
  PRINT "Done"
  KillFile "FIXMULTI.EXE"


