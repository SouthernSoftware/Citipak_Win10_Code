DEFINT A-Z

  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS

  DIM SHARED ARCustRec(1) AS ARCustRecType
  ARCustRecLen = LEN(ARCustRec(1))

  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  FOR zz = 1 TO NumOfArRecs
    GET ARFile, zz, ARCustRec(1)
    ARCustRec(1).IssueLicense = "Y"
    PUT ARFile, zz, ARCustRec(1)
  NEXT
  CLOSE


