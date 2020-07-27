DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

Date1 = Date2Num("07-14-1997")
Date2 = Date2Num("07-14-1998")

Date3 = Date2Num("08-11-1999")
Date4 = Date2Num("08-11-1998")


'$INCLUDE: 'PRTRANS.Bi'

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)

 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSN.DAT" FOR RANDOM AS #2 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen

 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec
   SELECT CASE Cnt
   CASE 229, 386, 414
     Found = Found + 1
   CASE ELSE
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE
 
 PRINT "Found:"; Found

