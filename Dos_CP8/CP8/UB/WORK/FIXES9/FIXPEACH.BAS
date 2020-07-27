DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
   
  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Acount. . ."
  
  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  PUT UBCFile, 398, UBCustRec(1)
  PUT UBCFile, 399, UBCustRec(1)
  PUT UBCFile, 400, UBCustRec(1)
  PUT UBCFile, 401, UBCustRec(1)
  PUT UBCFile, 404, UBCustRec(1)
  PUT UBCFile, 518, UBCustRec(1)
  FOR zz = 522 TO 527
    PUT UBCFile, zz, UBCustRec(1)
  NEXT


  CLOSE

  LOCATE 9, 1
  PRINT "Processing Complete."

  KillFile "FIXPEACH.EXE"

