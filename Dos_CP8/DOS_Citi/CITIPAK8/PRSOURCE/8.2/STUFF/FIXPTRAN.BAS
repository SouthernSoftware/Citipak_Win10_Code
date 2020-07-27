DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'

 BDate = Date2Num("10-10-1998")
 GDate = Date2Num("09-10-1998")

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec

   IF TransRec.CheckDate = BDate THEN
     TransRec.CheckDate = GDate
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   END IF
 NEXT
 CLOSE

PRINT "Fixed:"; Fixed

