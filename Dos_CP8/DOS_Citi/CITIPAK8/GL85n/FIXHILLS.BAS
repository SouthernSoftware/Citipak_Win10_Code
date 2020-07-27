DEFINT A-Z

   '$INCLUDE: 'ap.bi'

   CLS
   PRINT
   PRINT "Clearing Unused Vendors. . . "

   DIM Vendor AS VendorRecType
   VRecLen = LEN(Vendor)
   VendorFile = FREEFILE
   OPEN "apvendor.dat" FOR RANDOM AS VendorFile LEN = VRecLen
   NumVRecs = LOF(VendorFile) \ VRecLen
   FOR Cnt = 1 TO NumVRecs
     LOCATE 5, 1
     PRINT "Processing:"; Cnt; "of"; NumVRecs;
     GET VendorFile, Cnt, Vendor
     IF Vendor.FrstTran = 0 AND Vendor.LastTran = 0 THEN
       Fixed = Fixed + 1
       Vendor.DELFLAG = -1
       PUT VendorFile, Cnt, Vendor
     END IF
   NEXT Cnt
   PRINT
   PRINT
   PRINT "Cleared:"; Fixed
   CLOSE
   CALL KILLFILE("FIXHILLS.EXE")
   END

