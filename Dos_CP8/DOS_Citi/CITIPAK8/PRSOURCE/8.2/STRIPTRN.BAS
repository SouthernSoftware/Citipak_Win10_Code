DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

 'BadDate = Date2Num%("12-14-2000")

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
   SELECT CASE Cnt&
   CASE 766, 767, 768, 769
     Fixed = Fixed + 1
   CASE ELSE
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE
LOCATE 7, 1
PRINT "Removed:"; Fixed
KillFile "Striptrn.exe"
'SHELL "prutil1 relink 2000"

'JumpHere:
'
' OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
' NumOfRec& = LOF(1) \ TranLen
' FOR Cnt& = 1 TO NumOfRec&
'   LOCATE 4, 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
'   GET #1, Cnt&, TransRec
'   SELECT CASE TransRec.CheckNum
'   CASE 11223 TO 11232
'     STOP
'   END SELECT
' NEXT
'

