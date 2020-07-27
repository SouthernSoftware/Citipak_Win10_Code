DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)

'$INCLUDE: 'PRTRANS.Bi'
 
 L1Date = Date2Num("01-29-1996")
 T1Date = Date2Num("01-29-1998")

 L2Date = Date2Num("01-02-1987")
 T2Date = Date2Num("01-02-1998")

 L3Date = Date2Num("01-08-1985")
 T3Date = Date2Num("01-08-1998")
 CLS
 PRINT
 PRINT
 PRINT "Correcting Transaction Dates. . ."

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen
 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec
   SELECT CASE TransRec.CheckDate
   CASE L1Date
     TransRec.CheckDate = T1Date
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   CASE L2Date
     TransRec.CheckDate = T2Date
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   CASE L3Date
     TransRec.CheckDate = T3Date
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   END SELECT

 NEXT
 CLOSE
 PRINT
 PRINT "Corrected:"; Fixed

