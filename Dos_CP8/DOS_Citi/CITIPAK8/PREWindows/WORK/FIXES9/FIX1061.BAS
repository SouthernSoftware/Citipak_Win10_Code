DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Unsetting Prenote Flag. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  GET UBCFile, 1061, UBCustRec(1)
  UBCustRec(1).PreNoteFlag = 0
  PUT UBCFile, 1061, UBCustRec(1)
  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIX1061.EXE"

