DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

 FDate = Date2Num("03-30-2001")
' TDate = Date2Num("12-03-2001")
 
 CLS
 PRINT "Removing Transactions. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)

 KillFile "PRDATA\PRTRANSH.OLD"
 NAME "PRDATA\PRTRANSH.DAT" AS "PRDATA\PRTRANSH.OLD"

 OPEN "PRDATA\PRTRANSH.OLD" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #2 LEN = TranLen

 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   SELECT CASE TransRec.CheckDate
   CASE FDate
     Fixed = Fixed + 1
   CASE ELSE
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Removed:"; Fixed

KillFile "striptrn.exe"
'SHELL "prutil9 relink 2000"

