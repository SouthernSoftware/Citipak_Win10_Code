DEFINT A-Z

DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

'$INCLUDE: 'PRTRANS.Bi'

 TDate = Date2Num("09-05-2000")

 CLS
 PRINT "Correcting Transaction Date. . ."
 DIM TransRec AS TransRecType
 TranLen = LEN(TransRec)
 OPEN "PRDATA\PRTRANSH.DAT" FOR RANDOM AS #1 LEN = TranLen
 NumOfRec& = LOF(1) \ TranLen
 GET #1, 225, TransRec
 TransRec.CheckDate = TDate
 Fixed = Fixed + 1
 PUT #1, 225, TransRec
 CLOSE
  LOCATE 7, 1
  PRINT "Done."
  CALL KILLFile("FIXLANSI.EXE")

