DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

Date1 = Date2Num("12-23-1998")
Date2 = Date2Num("12-23-1999")

'Date3 = Date2Num("08-11-1999")
'Date4 = Date2Num("08-11-1998")


'$INCLUDE: 'PRTRANS.Bi'

 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec = LOF(1) \ TranLen

 FOR Cnt = 1 TO NumOfRec
   GET #1, Cnt, TransRec

   SELECT CASE TransRec.CheckDate
   CASE Date1
     TransRec.CheckDate = Date2
     Fixed = Fixed + 1
     PUT #1, Cnt, TransRec
   END SELECT
 NEXT
 CLOSE
 
 PRINT "Corrected:"; Fixed
 KillFile "fixtdate.exe"

