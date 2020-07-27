DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

'$INCLUDE: 'PRTRANS.Bi'

 FDate = Date2Num("11-19-1998")
 
 CLS
 PRINT "Adjusting Transaction Gross. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF TransRec.CheckDate = FDate AND TransRec.TotRegWage = 0 THEN
     TransRec.FedGrossPay = TransRec.FedGrossPay - TransRec.RetireAmt
     TransRec.StaGrossPay = TransRec.StaGrossPay - TransRec.RetireAmt
     Fixed = Fixed + 1
     PUT #1, Cnt&, TransRec
   END IF
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Corrected:"; Fixed
KILL "FIXGROSS.EXE"

