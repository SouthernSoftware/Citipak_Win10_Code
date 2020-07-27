DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB killFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

' FDate = Date2Num("08-31-1997")
 TDate = Date2Num("06-25-1999")
 
 CLS
 PRINT "Moving Deductions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)

 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen

 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF TransRec.CheckDate >= TDate THEN
     Fixed = Fixed + 1
     TransRec.DAMT(12) = TransRec.DAMT(9)
     TransRec.DAMT(9) = 0
     TransRec.DAMT(11) = TransRec.DAMT(8)
     TransRec.DAMT(8) = 0
     TransRec.DAMT(10) = TransRec.DAMT(7)
     TransRec.DAMT(7) = 0
     TransRec.DAMT(9) = TransRec.DAMT(6)
     TransRec.DAMT(6) = 0
     TransRec.DAMT(8) = TransRec.DAMT(5)
     TransRec.DAMT(5) = 0
     TransRec.DAMT(6) = TransRec.DAMT(4)
     TransRec.DAMT(4) = 0
     TransRec.DAMT(5) = TransRec.DAMT(3)
     TransRec.DAMT(3) = 0
     TransRec.DAMT(3) = TransRec.DAMT(2)
     TransRec.DAMT(2) = 0
     PUT #1, Cnt&, TransRec
   END IF
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Fixed:"; Fixed
killFile "FIXDEDS.EXE"

