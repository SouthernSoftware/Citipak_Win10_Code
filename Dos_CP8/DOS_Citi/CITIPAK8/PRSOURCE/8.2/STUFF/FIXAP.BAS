DEFINT A-Z

'$INCLUDE: 'PRTRANS.Bi'

 CLS
 PRINT "Correcting Invalid Entries. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec& = LOF(1) / TranLen
 FOR Cnt& = 1 TO NumOfRec&
   GET #1, Cnt&, TransRec
   SELECT CASE TransRec.CHECKNUM
   CASE 2562
     FOR zz = 1 TO 12
       IF TransRec.DAMT(zz) < 0 THEN
         TransRec.DAMT(zz) = 0
         Fixed = Fixed + 1
       END IF
     NEXT
     PUT #1, Cnt&, TransRec
   CASE 2523
     IF TransRec.RetireAmt < 0 THEN
       TransRec.RetireAmt = 0
       Fixed = Fixed + 1
     END IF
     PUT #1, Cnt&, TransRec
   CASE 2569
     IF TransRec.TotOTWage < 0 THEN
       TransRec.TotOTWage = 0
       Fixed = Fixed + 1
     END IF
     PUT #1, Cnt&, TransRec

   CASE 2579
     FOR zz = 1 TO 12
       IF TransRec.DAMT(zz) < 0 THEN
         TransRec.DAMT(zz) = 0
         Fixed = Fixed + 1
       END IF
     NEXT

     PUT #1, Cnt&, TransRec
   END SELECT
 NEXT
 CLOSE

PRINT
PRINT Fixed; "Transaction Correct."
PRINT
PRINT "Done"

