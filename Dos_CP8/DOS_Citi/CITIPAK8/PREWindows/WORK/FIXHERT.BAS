DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Book Numbers. . ."
  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  NumCust& = LOF(UBCFile) \ UBCustRecLen
  
  FOR CCnt& = 1 TO NumCust&
    GET UBCFile, CCnt&, UBCustRec(1)
    Book$ = QPTrim$(UBCustRec(1).Book)
    IF Book$ = "2" THEN
      UBCustRec(1).Book = "02"
      Fixed = Fixed + 1
      PUT UBCFile, CCnt&, UBCustRec(1)
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
PRINT "Fixed: "; Fixed
PRINT "Done."
CALL KillFile("fixHert.exe")

