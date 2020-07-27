DEFINT A-Z

'$INCLUDE: 'ap.bi'
  CLS
  LOCATE 3, 1
  PRINT "Swapping State and County Codes. . ."

  DIM Vendor AS VendorRecType
  VRecLen = LEN(Vendor)
  VendorFile = FREEFILE
  OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
  NumVRecs = LOF(VendorFile) \ VRecLen
  FOR Cnt = 1 TO NumVRecs
    LOCATE 6, 1: PRINT "Processing:"; Cnt; " of"; NumVRecs;
    GET VendorFile, Cnt, Vendor
    SWAP Vendor.CoCode, Vendor.StCode
    PUT VendorFile, Cnt, Vendor
  NEXT
  CLOSE
  LOCATE 12, 1
  PRINT "Processing Complete."
  KILL "SWAPCODE.EXE"
  END




