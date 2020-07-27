DEFINT A-Z

'$INCLUDE: 'PRTRANS.Bi'
 
 CLS
 PRINT "Correcting Transaction. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.NEW" FOR RANDOM AS #2 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   SELECT CASE Cnt
   CASE 42, 47, 110, 115, 161, 166, 234, 239, 284, 289, 335, 340, 386, 391, 436, 441, 465, 470, 496, 501
     D = D + 1
   CASE ELSE
     GET #1, Cnt, TransRec
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE
PRINT D
PRINT
PRINT
PRINT "Done"

