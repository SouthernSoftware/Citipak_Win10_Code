DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

' KillFile "PRDATA\PRTRANSH.NEW"
' KillFile "PRDATA\PRTRANSH.OLD"
' NAME "PRDATA\PRTRANSH.DAT" AS "PRDATA\PRTRANSH.OLD"

' LDate = Date2Num("12-17-1997")
' HDate = Date2Num("12-30-1997")
 
 CLS
 PRINT "Correcting Transactions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
' OPEN "PRDATA\PRTRANSH.OLD" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF TransRec.RetireAmt < -1000000 THEN
     TransRec.RetireAmt = 0
     Fixed = Fixed + 1
     PUT #1, Cnt&, TransRec
   END IF
 NEXT
 CLOSE
 LOCATE 7, 1
 PRINT "Fixed:"; Fixed

