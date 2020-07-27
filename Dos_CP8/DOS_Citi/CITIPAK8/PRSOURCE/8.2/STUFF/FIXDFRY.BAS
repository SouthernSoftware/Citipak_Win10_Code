DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'

 TDate = Date2Num("10-10-1997")
 
 CLS
 PRINT "Removing Transactions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.ZZZ" FOR RANDOM AS #2 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF TransRec.CHECKDate = TDate THEN
     Fixed = Fixed + 1
   ELSE
     PUT #2, , TransRec
   END IF
 NEXT
 CLOSE
 LOCATE 7, 1
 KILL "PRDATA\PRTRANSH.DAT"
 NAME "PRDATA\PRTRANSH.ZZZ" AS "PRDATA\PRTRANSH.DAT"
 PRINT "Removed:"; Fixed

