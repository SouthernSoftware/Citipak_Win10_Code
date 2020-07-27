DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

'$INCLUDE: 'PRTRANS.Bi'

 FDate = Date2Num("12-31-1979")
 PDate = Date2Num("12-31-1998")
 TDate = Date2Num("11-19-1998")
 
 CLS
 PRINT "Adjusting Transactions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF (TransRec.CheckDate = TDate) THEN
     IF TransRec.RegHrsPaid = 0 THEN
     Fixed = Fixed + 1
     TransRec.FedGrossPay = TransRec.GrossPay
     TransRec.StaGrossPay = TransRec.GrossPay
     'TransRec.CheckDate = TDate
     PUT #1, Cnt&, TransRec
     END IF
   END IF
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Corrected:"; Fixed
KILL "FIXMCCOR.EXE"

