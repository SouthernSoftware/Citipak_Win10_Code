DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

Date1 = Date2Num("03-15-1999")
Date2 = Date2Num("02-03-1999")

'Date3 = Date2Num("08-11-1999")
'Date4 = Date2Num("08-11-1998")


'$INCLUDE: 'PRTRANS.Bi'

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)

 NAME "PRDATA\PRTRANSH.DAT" AS "prdata\oldhist.trn"

 OPEN "prdata\oldhist.trn" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen

 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #2 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen

 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec
   SELECT CASE Cnt
   CASE 280, 276, 277, 273, 281, 270, 275, 274, 272, 279, 278, 271
   CASE ELSE
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE

 
 PRINT "Done"

