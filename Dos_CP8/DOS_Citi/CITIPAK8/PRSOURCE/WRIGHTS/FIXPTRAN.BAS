DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'



 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec
   IF TransRec.FedGrossPay < 0 THEN
     TransRec.FedGrossPay = TransRec.GROSSPAY
   END IF
   IF TransRec.StaGrossPay < 0 THEN
     TransRec.StaGrossPay = TransRec.GROSSPAY
   END IF

   PUT #1, Cnt, TransRec
 NEXT
 CLOSE


