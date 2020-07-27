DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'

 
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec
    IF TransRec.SocGrossPay <= 0 THEN
      STOP
    END IF
    IF TransRec.MedGrossPay <= 0 THEN
      STOP
    END IF
    IF TransRec.FedGrossPay <= 0 THEN
      STOP
    END IF
    IF TransRec.StaGrossPay <= 0 THEN
      STOP
    END IF
    IF TransRec.RetGrossPay <= 0 THEN
      STOP
    END IF

'   IF TransRec.FedTaxAmt < -500000 THEN
'     Fixed = Fixed + 1
'     TransRec.FedTaxAmt = 0
'   END IF
'   IF TransRec.StaTaxAmt < -500000 THEN
'     TransRec.StaTaxAmt = 0
'   END IF
' '  PUT #1, Cnt, TransRec
 NEXT
 CLOSE

'PRINT Fixed

