DEFINT A-Z
'$INCLUDE: 'ap.bi'
  CLS
  PRINT "Correcting Vendor. . . "

  DIM Vendor AS VendorRecType
  VRecLen = LEN(Vendor)
  VendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM SHARED AS VendorFile LEN = VRecLen
  NumVRecs = LOF(VendorFile) \ VRecLen
  FOR Cnt = 1 TO NumVRecs
  GET VendorFile, Cnt, Vendor
  IF INSTR(Vendor.VNum, "'M") > 0 AND INSTR(Vendor.VName, "'M") > 0 THEN
    LSET Vendor.VName = "MERICAN MUFFLER"
    LSET Vendor.VNum = "MERMUFF"
    Vendor.DelFlag = 0
    PUT VendorFile, Cnt, Vendor
  END IF
  NEXT
  CLOSE
  LOCATE 5, 1
  PRINT "Done"
  PRINT
  PRINT "Be sure to relink your apledger!!!!!!!!!!!!!"
  KILL "fixmer.exe"

