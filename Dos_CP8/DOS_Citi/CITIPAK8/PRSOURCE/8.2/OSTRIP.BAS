DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

' FDate = Date2Num("08-31-1997")
 TDate = Date2Num("06-25-1999")
 
 CLS
 PRINT "Moving Deductions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)

 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen

 OPEN "PRDATA\PRTRANSH.NEW" FOR RANDOM AS #2 LEN = TranLen

 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;

   IF Cnt& = 4920 OR Cnt& = 4922 OR Cnt& = 4925 THEN
     GET #1, Cnt&, TransRec
     PUT #2, , TransRec
   END IF
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Fixed:"; Fixed

