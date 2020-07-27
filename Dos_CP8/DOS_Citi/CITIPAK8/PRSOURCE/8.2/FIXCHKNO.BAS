DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB KillFile (FileName$)

'$INCLUDE: 'PRTRANS.Bi'

 CLS
 PRINT "Correcting Transaction. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 GET #1, 1705, TransRec
 TransRec.CheckNum = 6515
 PUT #1, 1705, TransRec
 CLOSE
 LOCATE 7, 1
 PRINT "Done"
 KillFile "fixchkno.exe"

