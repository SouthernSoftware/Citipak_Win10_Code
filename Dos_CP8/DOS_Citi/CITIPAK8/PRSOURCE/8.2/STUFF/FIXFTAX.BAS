DEFINT A-Z

'$INCLUDE: 'PRTRANS.Bi'
 
 CLS
 PRINT "Correcting Transaction. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 GET #1, 8651, TransRec
 TransRec.FedTaxAmt = 0
 PUT #1, 8651, TransRec
 CLOSE

PRINT
PRINT
PRINT "Done"

