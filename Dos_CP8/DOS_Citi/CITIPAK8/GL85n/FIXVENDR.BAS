DEFINT A-Z
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'ap.bi'
 CLS
 PRINT "Undeleting JERRY"

   DIM Vendor AS VendorRecType
   VRecLen = LEN(Vendor)
   VendorFile = FREEFILE
   OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS VendorFile LEN = VRecLen
   GET VendorFile, 444, Vendor
   Vendor.VNAME = "JERRY (I WAS DELETED)"
   Vendor.DELFLAG = 0
   PUT VendorFile, 444, Vendor
   CLOSE
   PRINT
   PRINT
   PRINT
   PRINT "Done."
   KillFile "FIXVENDR.EXE"

