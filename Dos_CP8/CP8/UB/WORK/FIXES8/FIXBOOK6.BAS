DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Book 06. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  FOR Cnt& = 1 TO CNumOfRecs&
    GET UBCFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Book = "06" THEN
      Seq$ = RIGHT$(UBCustRec(1).SeqNumb, 4)
      Seq$ = Seq$ + "00"
      UBCustRec(1).SeqNumb = Seq$
      PUT UBCFile, Cnt&, UBCustRec(1)
    END IF
  NEXT

LOCATE 9, 1
PRINT "Processing Complete."

