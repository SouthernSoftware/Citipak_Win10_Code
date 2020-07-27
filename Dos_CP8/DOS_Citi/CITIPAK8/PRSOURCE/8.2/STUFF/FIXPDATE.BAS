DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'

 BDate = Date2Num("12-22-1998")
 GDate = Date2Num("10-22-1998")

 CLS
 PRINT "Correcting transaction dates. . ."

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   LOCATE 3, 1
   PRINT "Processing:"; Cnt; " of"; NumOfRec;
   GET #1, Cnt, TransRec

   IF TransRec.CheckDate = BDate THEN
     TransRec.CheckDate = GDate
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   END IF
 NEXT
 CLOSE

PRINT
PRINT
PRINT "Done fixed:"; Fixed

