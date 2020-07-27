DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'
 TDate = Date2Num("01-01-1900")

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 GET #1, 23, TransRec
 TransRec.CheckDate = TDate
 TransRec.PostDate = TDate
 Fixed = Fixed + 1
 PUT #1, 23, TransRec
 CLOSE

PRINT "Done"

