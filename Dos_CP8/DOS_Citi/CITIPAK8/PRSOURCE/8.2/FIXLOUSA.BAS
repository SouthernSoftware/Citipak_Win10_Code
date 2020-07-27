DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

  BadDate = Date2Num("12-31-1979")
  GoodDate = Date2Num("06-30-2000")

 StartNo = 4984
 CLS
 PRINT "Correcting Transaction. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumTr& = LOF(1) \ TranLen

 FOR Cnt& = 1720 TO NumTr&
   GET #1, Cnt&, TransRec
   IF TransRec.CheckDate = BadDate THEN
     Fixed = Fixed + 1
     SELECT CASE Fixed
     CASE 1
       TransRec.CHECKNUM = 4985
     CASE 2
       TransRec.CHECKNUM = 4991
     CASE 3
       TransRec.CHECKNUM = 4986
     CASE 4
       TransRec.CHECKNUM = 4987
     CASE 5
       TransRec.CHECKNUM = 4988
     CASE 6
       TransRec.CHECKNUM = 4992
     CASE 7
       TransRec.CHECKNUM = 4996
     CASE 8
       TransRec.CHECKNUM = -1
     CASE 9
       TransRec.CHECKNUM = -1
     CASE 10
       TransRec.CHECKNUM = -1
     CASE 11
       TransRec.CHECKNUM = -1
     CASE 12
       TransRec.CHECKNUM = 4998
     CASE 13
       TransRec.CHECKNUM = 4990
     CASE 14
       TransRec.CHECKNUM = 4999
     CASE 15
       TransRec.CHECKNUM = -1
     CASE 16
       TransRec.CHECKNUM = 5001
     CASE 17
       TransRec.CHECKNUM = 5002
     CASE 18
       TransRec.CHECKNUM = 5003
     CASE 19
       TransRec.CHECKNUM = 4989
     END SELECT
     TransRec.CheckDate = GoodDate
     PUT #1, Cnt&, TransRec
   END IF
 NEXT
 
 CLOSE
 PRINT
 PRINT
 PRINT
 PRINT "Done"
 PRINT
 PRINT "Fixed:"; Fixed
 PRINT
 PRINT

 KillFile "fixlousa.exe"

 PRINT "Removing Transactions. . ."
 PRINT
 PRINT
 
 KillFile "PRDATA\PRTRANSH.OLD"
 NAME "PRDATA\PRTRANSH.DAT" AS "PRDATA\PRTRANSH.OLD"

 OPEN "PRDATA\PRTRANSH.OLD" FOR RANDOM AS #1 LEN = TranLen
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #2 LEN = TranLen

 NumOfRec& = LOF(1) \ TranLen
 FOR Cnt& = 1 TO NumOfRec&
   LOCATE , 1: PRINT " Scanning:"; Cnt&; " of"; NumOfRec&;
   GET #1, Cnt&, TransRec
   SELECT CASE TransRec.CHECKNUM
   CASE -1
     Removed = Removed + 1
   CASE ELSE
     PUT #2, , TransRec
   END SELECT
 NEXT
 CLOSE
PRINT
PRINT
PRINT "Removed:"; Removed

