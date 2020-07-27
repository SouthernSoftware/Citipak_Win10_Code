DEFINT A-Z

'$INCLUDE: 'ap.bi'

DIM Vendor  AS VendorRecType

VLen = LEN(Vendor)

OPEN "APVENDOR.dat" FOR RANDOM AS #1 LEN = VLen

   FOR Cnt = 1 TO LOF(1) \ VLen
     GET #1, Cnt, Vendor
     Vendor.VNum = UCASE$(Vendor.VNum)
     PUT #1, Cnt, Vendor
 NEXT
 CLOSE
call killfile("fixem.exe")

