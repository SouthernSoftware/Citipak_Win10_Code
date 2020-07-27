DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

 KillFile "PRDATA\PRTRANSH.NEW"
 KillFile "PRDATA\PRTRANSH.OLD"
 NAME "PRDATA\PRTRANSH.DAT" AS "PRDATA\PRTRANSH.OLD"

 LDate = Date2Num("12-17-1997")
 HDate = Date2Num("12-30-1997")
 
 CLS
 PRINT "Removing Transactions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.OLD" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #2 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   IF TransRec.CheckDate > LDate AND TransRec.CheckDate < HDate THEN
     Fixed = Fixed + 1
   ELSE
     PUT #2, , TransRec
   END IF
 NEXT
 CLOSE
 LOCATE 7, 1
 PRINT "Removed:"; Fixed

